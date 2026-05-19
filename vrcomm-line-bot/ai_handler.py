"""
ai_handler.py — Entry point for LINE message AI processing.
Classifies intent then routes to the appropriate agent.
"""
import logging
from intent_router import classify_intent, route

logger = logging.getLogger(__name__)


def process_with_ai(user_name: str, user_id: str, message: str,
                    source_type: str = "user", history: list = None) -> str:
    """
    Main entry point called by app.py for every incoming LINE text message.

    1. Classify intent (fast Haiku call)
    2. Route to the matching agent
    3. Return reply string
    """
    if history is None:
        history = []

    # Map LINE source_type to a source label
    source = "line"
    if source_type == "group":
        source = "line_group"
    elif source_type == "room":
        source = "line_room"

    intent = classify_intent(message, history)

    return route(
        intent=intent,
        message=message,
        user_name=user_name,
        user_id=user_id,
        source=source,
        history=history,
    )
