"""
Event notification utilities for cross-component status updates.
"""

from typing import Callable, List, Dict, Any


class EventBus:
    """Lightweight in-memory event bus."""

    def __init__(self):
        self._subscribers: Dict[str, List[Callable[..., Any]]] = {}

    def subscribe(self, event_name: str, callback: Callable[..., Any]):
        if event_name not in self._subscribers:
            self._subscribers[event_name] = []
        self._subscribers[event_name].append(callback)

    def unsubscribe(self, event_name: str, callback: Callable[..., Any]):
        if event_name in self._subscribers and callback in self._subscribers[event_name]:
            self._subscribers[event_name].remove(callback)

    def emit(self, event_name: str, *args, **kwargs):
        if event_name in self._subscribers:
            for cb in self._subscribers[event_name]:
                try:
                    cb(*args, **kwargs)
                except Exception as err:
                    print(f"[event_bus] Error invoking handler for {event_name}: {err}")


# Global app event bus
app_events = EventBus()
