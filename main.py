#!/usr/bin/env python3
"""
OpenCom media -> RPC bridge
Supports:
- KDE/Linux via MPRIS over D-Bus
- Windows via GlobalSystemMediaTransportControlsSessionManager

RPC target:
- GET    /rpc/health
- POST   /rpc/activity
- DELETE /rpc/activity

Install:
  Common:
    pip install requests

  Linux (KDE / MPRIS):
    pip install dbus-next

  Windows:
    pip install winsdk

Run:
  python main.py
"""

from __future__ import annotations

import asyncio
import platform
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Any

# Allow running the repo script without installing the package first.
SRC_DIR = Path(__file__).resolve().parent / "src"
if SRC_DIR.exists():
    sys.path.insert(0, str(SRC_DIR))

from opencom_rpc import (
    Activity,
    ActivityButton,
    DEFAULT_MEDIA_POLL_SECONDS,
    OpenComRPCClient,
)

POLL_SECONDS = DEFAULT_MEDIA_POLL_SECONDS
PROGRESS_BAR_WIDTH = 12
_LINUX_BUS: Any | None = None


@dataclass
class MediaState:
    source: str
    title: str
    artist: str
    album: str
    is_playing: bool
    duration_ms: Optional[int] = None
    position_ms: Optional[int] = None
    track_url: Optional[str] = None
    art_url: Optional[str] = None


def fmt_ms(ms: Optional[int]) -> str:
    if ms is None:
        return "?"
    total = ms // 1000
    m, s = divmod(total, 60)
    return f"{m}:{s:02d}"


def build_progress_bar(position_ms: Optional[int], duration_ms: Optional[int], width: int = PROGRESS_BAR_WIDTH) -> str:
    if position_ms is None or duration_ms is None or duration_ms <= 0:
        return ""

    clamped_position = max(0, min(position_ms, duration_ms))
    filled = round((clamped_position / duration_ms) * width)
    filled = max(0, min(filled, width))
    return f"[{'=' * filled}{'-' * (width - filled)}]"


def build_progress_text(media: MediaState) -> str:
    if media.position_ms is None or media.duration_ms is None or media.duration_ms <= 0:
        return ""

    progress_bar = build_progress_bar(media.position_ms, media.duration_ms)
    return f"{fmt_ms(media.position_ms)}/{fmt_ms(media.duration_ms)} {progress_bar}"


def build_state_text(media: MediaState, max_len: int = 128) -> Optional[str]:
    artist_album = media.artist
    if media.album:
        artist_album = f"{media.artist} • {media.album}" if media.artist else media.album

    progress_text = build_progress_text(media)
    if not progress_text:
        return artist_album[:max_len] if artist_album else None

    if not artist_album:
        return progress_text[:max_len]

    separator = " | "
    full_state = f"{artist_album}{separator}{progress_text}"
    if len(full_state) <= max_len:
        return full_state

    reserved = len(separator) + len(progress_text)
    if reserved >= max_len:
        return progress_text[:max_len]

    return f"{artist_album[: max_len - reserved]}{separator}{progress_text}"


def progress_bucket(media: MediaState) -> Optional[int]:
    if media.position_ms is None or media.duration_ms is None or media.duration_ms <= 0:
        return None

    interval_ms = max(1000, int(POLL_SECONDS * 1000))
    return media.position_ms // interval_ms


async def linux_get_media() -> Optional[MediaState]:
    try:
        from dbus_next.aio import MessageBus
    except ImportError:
        print("Missing dependency: dbus-next")
        print("Install with: pip install dbus-next")
        return None

    global _LINUX_BUS

    try:
        if _LINUX_BUS is None or not _LINUX_BUS.connected:
            _LINUX_BUS = await MessageBus().connect()

        bus = _LINUX_BUS
    except Exception as exc:
        print(f"[linux] failed to connect to D-Bus: {exc}")
        close_linux_bus()
        return None

    try:
        # Find MPRIS players
        introspection = await bus.introspect("org.freedesktop.DBus", "/org/freedesktop/DBus")
        obj = bus.get_proxy_object("org.freedesktop.DBus", "/org/freedesktop/DBus", introspection)
        iface = obj.get_interface("org.freedesktop.DBus")
        names = await iface.call_list_names()
    except Exception as exc:
        print(f"[linux] failed to read media players: {exc}")
        close_linux_bus()
        return None

    players = sorted(name for name in names if name.startswith("org.mpris.MediaPlayer2."))
    if not players:
        return None

    best: Optional[MediaState] = None

    for player in players:
        try:
            intr = await bus.introspect(player, "/org/mpris/MediaPlayer2")
            pobj = bus.get_proxy_object(player, "/org/mpris/MediaPlayer2", intr)
            props = pobj.get_interface("org.freedesktop.DBus.Properties")

            playback_status_var = await props.call_get("org.mpris.MediaPlayer2.Player", "PlaybackStatus")
            metadata_var = await props.call_get("org.mpris.MediaPlayer2.Player", "Metadata")
            position_var = await props.call_get("org.mpris.MediaPlayer2.Player", "Position")

            playback_status = str(playback_status_var.value)
            metadata = metadata_var.value if hasattr(metadata_var, "value") else {}
            position_us = position_var.value if hasattr(position_var, "value") else 0

            title = _variant_to_str(metadata.get("xesam:title"))
            artists = _variant_to_list(metadata.get("xesam:artist"))
            album = _variant_to_str(metadata.get("xesam:album"))
            art_url = _variant_to_str(metadata.get("mpris:artUrl"))
            track_url = _variant_to_str(metadata.get("xesam:url"))
            length_us = _variant_to_int(metadata.get("mpris:length"))

            state = MediaState(
                source=player.replace("org.mpris.MediaPlayer2.", "", 1),
                title=title or "Unknown title",
                artist=", ".join(artists) if artists else "Unknown artist",
                album=album or "",
                is_playing=(playback_status.lower() == "playing"),
                duration_ms=(length_us // 1000) if length_us is not None else None,
                position_ms=(position_us // 1000) if position_us is not None else None,
                track_url=track_url or None,
                art_url=art_url or None,
            )

            # Prefer actively playing sessions
            if state.is_playing:
                return state
            if best is None:
                best = state
        except Exception:
            continue

    return best


def close_linux_bus() -> None:
    global _LINUX_BUS

    if _LINUX_BUS is None:
        return

    try:
        _LINUX_BUS.disconnect()
    except Exception:
        pass
    finally:
        _LINUX_BUS = None


def _variant_to_str(value: Any) -> str:
    if value is None:
        return ""
    if hasattr(value, "value"):
        value = value.value
    if isinstance(value, str):
        return value
    return str(value)


def _variant_to_list(value: Any) -> list[str]:
    if value is None:
        return []
    if hasattr(value, "value"):
        value = value.value
    if isinstance(value, list):
        return [str(x) for x in value]
    return [str(value)]


def _variant_to_int(value: Any) -> Optional[int]:
    if value is None:
        return None
    if hasattr(value, "value"):
        value = value.value
    try:
        return int(value)
    except Exception:
        return None


async def windows_get_media() -> Optional[MediaState]:
    try:
        from winsdk.windows.media.control import GlobalSystemMediaTransportControlsSessionManager
    except ImportError:
        print("Missing dependency: winsdk")
        print("Install with: pip install winsdk")
        return None

    try:
        manager = await GlobalSystemMediaTransportControlsSessionManager.request_async()
        session = manager.get_current_session()
        if session is None:
            return None

        info = await session.try_get_media_properties_async()
        playback = session.get_playback_info()
        timeline = session.get_timeline_properties()

        artist = getattr(info, "artist", "") or "Unknown artist"
        title = getattr(info, "title", "") or "Unknown title"
        album = getattr(info, "album_title", "") or ""

        art_url = None
        track_url = None

        # Windows API often doesn't expose a public track URL here.
        # Thumbnail extraction is possible, but keeping this version simple and reliable.
        status = getattr(playback, "playback_status", None)
        is_playing = str(status).lower().endswith("playing")

        start = getattr(timeline, "start_time", None)
        end = getattr(timeline, "end_time", None)
        pos = getattr(timeline, "position", None)

        duration_ms = _td_to_ms(end) - _td_to_ms(start) if end is not None and start is not None else None
        position_ms = _td_to_ms(pos) if pos is not None else None

        source_app = session.source_app_user_model_id or "windows-media"

        return MediaState(
            source=source_app,
            title=title,
            artist=artist,
            album=album,
            is_playing=is_playing,
            duration_ms=duration_ms if duration_ms and duration_ms > 0 else None,
            position_ms=position_ms,
            track_url=track_url,
            art_url=art_url,
        )
    except Exception as exc:
        print(f"[windows] failed to read media: {exc}")
        return None


def _td_to_ms(value: Any) -> int:
    # winsdk TimeSpan usually exposes duration in 100ns ticks via .duration
    try:
        ticks = int(value.duration)
        return ticks // 10_000
    except Exception:
        try:
            return int(value.total_milliseconds())
        except Exception:
            return 0


def print_media(media: Optional[MediaState]) -> None:
    if media is None:
        print("[media] nothing active")
        return

    icon = "▶" if media.is_playing else "⏸"
    print(
        f"[media] {icon} {media.title} — {media.artist}"
        f" [{fmt_ms(media.position_ms)}/{fmt_ms(media.duration_ms)}]"
        f" ({media.source})"
    )


async def get_media_state() -> Optional[MediaState]:
    system = platform.system().lower()
    if system == "windows":
        return await windows_get_media()
    if system == "linux":
        return await linux_get_media()
    print(f"Unsupported platform: {platform.system()}")
    return None


def media_signature(media: Optional[MediaState]) -> Optional[tuple]:
    if media is None:
        return None
    return (
        media.source,
        media.title,
        media.artist,
        media.album,
        media.is_playing,
        media.duration_ms,
        progress_bucket(media) if media.is_playing else None,
        media.track_url,
        media.art_url,
    )


def media_to_activity(media: MediaState) -> Activity:
    start_timestamp = None
    end_timestamp = None
    if media.is_playing and media.position_ms is not None:
        now_ms = int(time.time() * 1000)
        start_timestamp = max(0, now_ms - media.position_ms)
        if media.duration_ms is not None and media.duration_ms > 0:
            end_timestamp = start_timestamp + media.duration_ms

    buttons = []
    if media.track_url:
        buttons.append(ActivityButton(label="Open Track", url=media.track_url))

    return Activity(
        name="Listening to music",
        details=media.title[:128],
        state=build_state_text(media),
        start_timestamp=start_timestamp,
        end_timestamp=end_timestamp,
        large_image_url=media.art_url,
        buttons=buttons[:2],
    )


def get_rpc_health(rpc: OpenComRPCClient) -> dict[str, Any]:
    try:
        health = rpc.health()
        return health if isinstance(health, dict) else {"ok": True, "response": health}
    except Exception as exc:
        return {"ok": False, "error": str(exc)}


def clear_rpc_activity(rpc: OpenComRPCClient) -> None:
    try:
        rpc.clear_activity()
    except Exception as exc:
        print(f"[rpc] clear failed: {exc}")


def post_rpc_activity(rpc: OpenComRPCClient, media: MediaState) -> None:
    try:
        rpc.set_activity(media_to_activity(media))
    except Exception as exc:
        print(f"[rpc] post failed: {exc}")


async def main() -> None:
    rpc = OpenComRPCClient()
    try:
        print(f"OpenCom RPC base: {rpc.base_url}")
        health = get_rpc_health(rpc)
        print(f"RPC health: {health}")

        last_sig = None
        last_clear = False

        while True:
            try:
                media = await get_media_state()
                sig = media_signature(media)

                if media is None:
                    if not last_clear:
                        print_media(None)
                        clear_rpc_activity(rpc)
                        last_clear = True
                    await asyncio.sleep(POLL_SECONDS)
                    continue

                print_media(media)

                if sig != last_sig:
                    post_rpc_activity(rpc, media)
                    last_sig = sig
                    last_clear = False

                # Optional: if paused, clear instead of showing paused
                # Uncomment these lines if you prefer that behavior:
                #
                # if not media.is_playing:
                #     rpc.clear_activity()
                #     last_clear = True

            except KeyboardInterrupt:
                print("\nExiting...")
                break
            except Exception as exc:
                print(f"[loop] error: {exc}")

            await asyncio.sleep(POLL_SECONDS)
    finally:
        close_linux_bus()
        rpc.close()


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        pass
