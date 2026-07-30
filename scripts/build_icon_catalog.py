#!/usr/bin/env python3
"""Build the shared IconAid catalog and PowerPoint vector renderer."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parent.parent
CATALOG_PATH = ROOT / "shared" / "iconaid" / "catalog.json"
VBA_PATH = ROOT / "apps" / "powerpoint" / "src" / "modIconAid.bas"


def line(x1: float, y1: float, x2: float, y2: float) -> dict[str, Any]:
    return {"kind": "line", "x1": x1, "y1": y1, "x2": x2, "y2": y2}


def rect(x: float, y: float, width: float, height: float, *, filled: bool = False) -> dict[str, Any]:
    return {"kind": "rect", "x": x, "y": y, "width": width, "height": height, "filled": filled}


def ellipse(x: float, y: float, width: float, height: float, *, filled: bool = False) -> dict[str, Any]:
    return {"kind": "ellipse", "x": x, "y": y, "width": width, "height": height, "filled": filled}


def polyline(*points: tuple[float, float], closed: bool = False) -> list[dict[str, Any]]:
    pairs = list(zip(points, points[1:]))
    if closed and len(points) > 2:
        pairs.append((points[-1], points[0]))
    return [line(start[0], start[1], end[0], end[1]) for start, end in pairs]


ICONS: list[dict[str, Any]] = [
    {
        "id": "analytics",
        "name": "Analytics",
        "category": "Business",
        "tags": ["chart", "metrics", "reporting", "growth", "data"],
        "primitives": [
            line(3, 20, 21, 20), line(3, 20, 3, 4),
            rect(6, 13, 3, 7), rect(11, 9, 3, 11), rect(16, 5, 3, 15),
        ],
    },
    {
        "id": "briefcase",
        "name": "Briefcase",
        "category": "Business",
        "tags": ["work", "consulting", "job", "portfolio", "company"],
        "primitives": [
            rect(3, 7, 18, 13), rect(8, 4, 8, 3),
            line(3, 12, 21, 12), rect(10.5, 10.5, 3, 3, filled=True),
        ],
    },
    {
        "id": "target",
        "name": "Target",
        "category": "Business",
        "tags": ["goal", "objective", "strategy", "focus", "kpi"],
        "primitives": [
            ellipse(3, 3, 18, 18), ellipse(7, 7, 10, 10), ellipse(10.5, 10.5, 3, 3, filled=True),
            line(13, 11, 21, 3), line(18, 3, 21, 3), line(21, 3, 21, 6),
        ],
    },
    {
        "id": "users",
        "name": "Team",
        "category": "Business",
        "tags": ["people", "team", "customer", "organization", "stakeholders"],
        "primitives": [
            ellipse(9, 4, 6, 6), ellipse(3, 7, 4, 4), ellipse(17, 7, 4, 4),
            line(6, 20, 6, 18), line(6, 18, 9, 14), line(9, 14, 15, 14),
            line(15, 14, 18, 18), line(18, 18, 18, 20), line(9, 20, 9, 18),
            line(9, 18, 11, 16), line(11, 16, 13, 16), line(13, 16, 15, 18), line(15, 18, 15, 20),
        ],
    },
    {
        "id": "cloud",
        "name": "Cloud",
        "category": "Technology",
        "tags": ["saas", "hosting", "internet", "platform", "infrastructure"],
        "primitives": [
            ellipse(3, 10, 7, 7), ellipse(7, 5, 10, 10), ellipse(14, 9, 7, 7),
            line(6.5, 17, 18, 17),
        ],
    },
    {
        "id": "database",
        "name": "Database",
        "category": "Technology",
        "tags": ["data", "storage", "sql", "warehouse", "server"],
        "primitives": [
            ellipse(4, 3, 16, 5), ellipse(4, 9.5, 16, 5), ellipse(4, 16, 16, 5),
            line(4, 5.5, 4, 18.5), line(20, 5.5, 20, 18.5),
        ],
    },
    {
        "id": "server",
        "name": "Server",
        "category": "Technology",
        "tags": ["hosting", "infrastructure", "rack", "compute", "backend"],
        "primitives": [
            rect(3, 3, 18, 5), rect(3, 9.5, 18, 5), rect(3, 16, 18, 5),
            ellipse(5, 5, 1.5, 1.5, filled=True), ellipse(5, 11.5, 1.5, 1.5, filled=True),
            ellipse(5, 18, 1.5, 1.5, filled=True), line(9, 5.75, 18.5, 5.75),
            line(9, 12.25, 18.5, 12.25), line(9, 18.75, 18.5, 18.75),
        ],
    },
    {
        "id": "laptop",
        "name": "Laptop",
        "category": "Technology",
        "tags": ["computer", "device", "software", "workstation", "screen"],
        "primitives": [
            rect(5, 4, 14, 11), line(3, 19, 21, 19), line(3, 19, 5, 15),
            line(21, 19, 19, 15), line(9, 17, 15, 17),
        ],
    },
    {
        "id": "mobile",
        "name": "Mobile",
        "category": "Technology",
        "tags": ["phone", "smartphone", "app", "device", "digital"],
        "primitives": [
            rect(7, 2.5, 10, 19), line(7, 6, 17, 6), line(7, 18, 17, 18),
            ellipse(11.25, 19, 1.5, 1.5, filled=True),
        ],
    },
    {
        "id": "chip",
        "name": "AI Chip",
        "category": "Technology",
        "tags": ["ai", "processor", "semiconductor", "machine learning", "compute"],
        "primitives": [
            rect(6, 6, 12, 12), rect(9, 9, 6, 6),
            line(3, 8, 6, 8), line(3, 12, 6, 12), line(3, 16, 6, 16),
            line(18, 8, 21, 8), line(18, 12, 21, 12), line(18, 16, 21, 16),
            line(8, 3, 8, 6), line(12, 3, 12, 6), line(16, 3, 16, 6),
            line(8, 18, 8, 21), line(12, 18, 12, 21), line(16, 18, 16, 21),
        ],
    },
    {
        "id": "network",
        "name": "Network",
        "category": "Technology",
        "tags": ["nodes", "connections", "architecture", "ecosystem", "distributed"],
        "primitives": [
            line(6, 6, 18, 6), line(6, 6, 12, 18), line(18, 6, 12, 18),
            ellipse(3, 3, 6, 6, filled=True), ellipse(15, 3, 6, 6, filled=True),
            ellipse(9, 15, 6, 6, filled=True),
        ],
    },
    {
        "id": "lock",
        "name": "Security",
        "category": "Technology",
        "tags": ["lock", "privacy", "cybersecurity", "protection", "compliance"],
        "primitives": [
            ellipse(7, 3, 10, 11), rect(5, 10, 14, 11),
            ellipse(11, 14, 2, 2, filled=True), line(12, 16, 12, 18),
        ],
    },
    {
        "id": "presentation",
        "name": "Presentation",
        "category": "Business",
        "tags": ["slides", "powerpoint", "google slides", "pitch", "deck", "meeting", "storytelling", "speaker"],
        "primitives": [
            rect(3, 3, 18, 13), line(12, 16, 12, 21), line(8, 21, 16, 21),
            rect(6, 10, 2.5, 3), rect(10.5, 7, 2.5, 6), rect(15, 5, 2.5, 8),
        ],
    },
    {
        "id": "calendar",
        "name": "Calendar",
        "category": "Business",
        "tags": ["date", "schedule", "planning", "timeline", "appointment", "roadmap", "milestone", "event"],
        "primitives": [
            rect(3, 5, 18, 16), line(3, 9, 21, 9), line(7, 3, 7, 7), line(17, 3, 17, 7),
            rect(6, 12, 3, 3, filled=True), rect(11, 12, 3, 3, filled=True), rect(16, 12, 3, 3, filled=True),
            rect(6, 17, 3, 2, filled=True), rect(11, 17, 3, 2, filled=True),
        ],
    },
    {
        "id": "clock",
        "name": "Clock",
        "category": "Business",
        "tags": ["time", "deadline", "duration", "schedule", "speed", "wait", "history", "productivity"],
        "primitives": [ellipse(3, 3, 18, 18), line(12, 6, 12, 12), line(12, 12, 17, 15)],
    },
    {
        "id": "lightbulb",
        "name": "Idea",
        "category": "Business",
        "tags": ["lightbulb", "innovation", "insight", "concept", "creative", "brainstorm", "solution", "inspiration"],
        "primitives": [
            ellipse(7, 3, 10, 12), line(9, 13, 10, 17), line(15, 13, 14, 17),
            line(10, 17, 14, 17), line(10, 20, 14, 20), line(12, 0.5, 12, 2),
            line(3.5, 5, 5, 6), line(19, 6, 20.5, 5),
        ],
    },
    {
        "id": "compass",
        "name": "Compass",
        "category": "Business",
        "tags": ["direction", "strategy", "navigation", "vision", "north star", "guidance", "orientation", "explore"],
        "primitives": [
            ellipse(3, 3, 18, 18), ellipse(10.5, 10.5, 3, 3, filled=True),
            *polyline((8, 16), (10.5, 10.5), (16, 8), (13.5, 13.5), (8, 16), closed=True),
        ],
    },
    {
        "id": "flag",
        "name": "Flag",
        "category": "Business",
        "tags": ["milestone", "goal", "achievement", "priority", "launch", "finish", "objective", "marker"],
        "primitives": [line(5, 3, 5, 21), *polyline((5, 4), (18, 4), (15, 9), (18, 14), (5, 14))],
    },
    {
        "id": "puzzle",
        "name": "Puzzle",
        "category": "Business",
        "tags": ["solution", "fit", "integration", "problem solving", "piece", "strategy", "complexity", "synergy"],
        "primitives": [
            rect(3, 3, 8, 8), rect(13, 3, 8, 8), rect(3, 13, 8, 8), rect(13, 13, 8, 8),
            ellipse(9, 5.5, 4, 3), ellipse(5.5, 9, 3, 4), ellipse(15.5, 9, 3, 4), ellipse(9, 15.5, 4, 3),
        ],
    },
    {
        "id": "handshake",
        "name": "Partnership",
        "category": "Business",
        "tags": ["handshake", "deal", "agreement", "collaboration", "partner", "trust", "alliance", "contract"],
        "primitives": [
            *polyline((2, 9), (6, 7), (10, 11), (12, 9), (15, 8), (22, 13)),
            *polyline((2, 15), (7, 18), (10, 15), (13, 18), (17, 14), (20, 15), (22, 13)),
            line(7, 18, 4, 16), line(17, 14, 14, 11),
        ],
    },
    {
        "id": "currency",
        "name": "Currency",
        "category": "Finance",
        "tags": ["money", "dollar", "cash", "revenue", "sales", "price", "cost", "financial", "usd"],
        "primitives": [
            ellipse(3, 3, 18, 18), line(12, 5, 12, 19),
            *polyline((16, 8), (14, 6.5), (10, 6.5), (8, 8), (10, 11), (14, 11), (16, 13), (14, 17), (10, 17), (8, 15)),
        ],
    },
    {
        "id": "percent",
        "name": "Percent",
        "category": "Finance",
        "tags": ["percentage", "ratio", "discount", "margin", "rate", "share", "growth", "conversion", "interest"],
        "primitives": [ellipse(4, 4, 5, 5), ellipse(15, 15, 5, 5), line(6, 19, 18, 5)],
    },
    {
        "id": "bank",
        "name": "Bank",
        "category": "Finance",
        "tags": ["institution", "finance", "capital", "treasury", "government", "investment", "funding", "loan"],
        "primitives": [
            *polyline((3, 8), (12, 3), (21, 8), closed=True), line(3, 9, 21, 9),
            rect(5, 10, 2.5, 8), rect(10.75, 10, 2.5, 8), rect(16.5, 10, 2.5, 8),
            line(3, 19, 21, 19), line(2, 21, 22, 21),
        ],
    },
    {
        "id": "wallet",
        "name": "Wallet",
        "category": "Finance",
        "tags": ["payment", "cash", "budget", "spend", "expense", "account", "money", "billing"],
        "primitives": [
            rect(3, 6, 18, 14), line(3, 9, 21, 9), rect(14, 11, 7, 5),
            ellipse(16, 12.75, 1.5, 1.5, filled=True),
        ],
    },
    {
        "id": "calculator",
        "name": "Calculator",
        "category": "Finance",
        "tags": ["calculate", "math", "accounting", "forecast", "budget", "estimate", "numbers", "finance"],
        "primitives": [
            rect(5, 2, 14, 20), rect(8, 5, 8, 4),
            *[ellipse(x, y, 1.75, 1.75, filled=True) for y in (12, 16.5) for x in (8, 11.25, 14.5)],
        ],
    },
    {
        "id": "invoice",
        "name": "Invoice",
        "category": "Finance",
        "tags": ["bill", "receipt", "document", "payment", "accounts payable", "accounts receivable", "transaction", "purchase"],
        "primitives": [
            rect(5, 2, 14, 20), line(8, 7, 16, 7), line(8, 11, 16, 11), line(8, 15, 13, 15),
            line(14, 17, 17, 17), line(15.5, 15.5, 15.5, 18.5),
        ],
    },
    {
        "id": "trend-up",
        "name": "Trend Up",
        "category": "Finance",
        "tags": ["growth", "increase", "revenue", "performance", "bullish", "forecast", "improvement", "upward"],
        "primitives": [
            line(3, 20, 21, 20), line(3, 20, 3, 4),
            *polyline((5, 16), (9, 12), (13, 14), (19, 6)), line(15, 6, 19, 6), line(19, 6, 19, 10),
        ],
    },
    {
        "id": "trend-down",
        "name": "Trend Down",
        "category": "Finance",
        "tags": ["decline", "decrease", "loss", "performance", "bearish", "forecast", "reduction", "downward"],
        "primitives": [
            line(3, 20, 21, 20), line(3, 20, 3, 4),
            *polyline((5, 7), (9, 11), (13, 9), (19, 17)), line(15, 17, 19, 17), line(19, 13, 19, 17),
        ],
    },
    {
        "id": "code",
        "name": "Code",
        "category": "Technology",
        "tags": ["software", "developer", "programming", "source", "engineering", "html", "script", "application"],
        "primitives": [
            *polyline((9, 6), (4, 12), (9, 18)), *polyline((15, 6), (20, 12), (15, 18)),
            line(14, 4, 10, 20),
        ],
    },
    {
        "id": "terminal",
        "name": "Terminal",
        "category": "Technology",
        "tags": ["command line", "cli", "shell", "console", "developer", "devops", "script", "code"],
        "primitives": [rect(3, 4, 18, 16), *polyline((7, 9), (10, 12), (7, 15)), line(12, 15, 17, 15)],
    },
    {
        "id": "api",
        "name": "API",
        "category": "Technology",
        "tags": ["integration", "interface", "endpoint", "service", "microservice", "connection", "backend", "system"],
        "primitives": [
            rect(8, 8, 8, 8), ellipse(2, 9, 4, 4, filled=True), ellipse(18, 3, 4, 4, filled=True),
            ellipse(18, 17, 4, 4, filled=True), line(6, 11, 8, 11), line(16, 10, 19, 7), line(16, 14, 19, 17),
        ],
    },
    {
        "id": "automation",
        "name": "Automation",
        "category": "Technology",
        "tags": ["workflow", "robotic process automation", "rpa", "gear", "efficiency", "orchestration", "automatic", "process"],
        "primitives": [
            ellipse(7, 7, 10, 10), ellipse(10.5, 10.5, 3, 3, filled=True),
            rect(10.5, 2, 3, 4), rect(10.5, 18, 3, 4), rect(2, 10.5, 4, 3), rect(18, 10.5, 4, 3),
            line(5, 5, 7.5, 7.5), line(16.5, 16.5, 19, 19), line(19, 5, 16.5, 7.5), line(7.5, 16.5, 5, 19),
        ],
    },
    {
        "id": "globe",
        "name": "Global",
        "category": "Technology",
        "tags": ["globe", "world", "international", "internet", "website", "geography", "market", "earth"],
        "primitives": [
            ellipse(3, 3, 18, 18), ellipse(8, 3, 8, 18), line(3, 12, 21, 12), line(5, 7, 19, 7), line(5, 17, 19, 17),
        ],
    },
    {
        "id": "wifi",
        "name": "Wireless",
        "category": "Technology",
        "tags": ["wifi", "wireless", "signal", "connectivity", "network", "internet", "radio", "online"],
        "primitives": [
            ellipse(10.5, 18, 3, 3, filled=True),
            line(12, 18, 6, 12), line(12, 18, 18, 12), line(12, 18, 3, 9), line(12, 18, 21, 9),
            line(12, 18, 1, 6), line(12, 18, 23, 6),
        ],
    },
    {
        "id": "layers",
        "name": "Layers",
        "category": "Technology",
        "tags": ["stack", "architecture", "platform", "technology stack", "components", "system", "infrastructure", "levels"],
        "primitives": [
            *polyline((3, 8), (12, 3), (21, 8), (12, 13), (3, 8), closed=True),
            *polyline((3, 12), (12, 17), (21, 12)), *polyline((3, 16), (12, 21), (21, 16)),
        ],
    },
    {
        "id": "webhook",
        "name": "Webhook",
        "category": "Technology",
        "tags": ["event", "callback", "integration", "trigger", "automation", "api", "notification", "endpoint"],
        "primitives": [
            ellipse(10, 2, 4, 4, filled=True), ellipse(3, 16, 4, 4, filled=True), ellipse(17, 16, 4, 4, filled=True),
            line(12, 6, 6, 16), line(12, 6, 18, 16), line(7, 18, 17, 18),
        ],
    },
    {
        "id": "factory",
        "name": "Factory",
        "category": "Operations",
        "tags": ["manufacturing", "industry", "plant", "production", "operations", "facility", "industrial", "output"],
        "primitives": [
            *polyline((3, 20), (3, 10), (8, 13), (8, 9), (13, 12), (13, 8), (18, 11), (18, 20)),
            rect(18, 4, 3, 16), line(3, 20, 21, 20), rect(6, 16, 3, 4), rect(12, 16, 3, 4),
        ],
    },
    {
        "id": "truck",
        "name": "Truck",
        "category": "Operations",
        "tags": ["logistics", "delivery", "shipping", "transport", "freight", "supply chain", "distribution", "vehicle"],
        "primitives": [
            rect(2, 7, 12, 10), rect(14, 11, 6, 6), *polyline((14, 11), (17, 7), (20, 11)),
            ellipse(5, 15, 4, 4), ellipse(15, 15, 4, 4), line(20, 17, 22, 17),
        ],
    },
    {
        "id": "package",
        "name": "Package",
        "category": "Operations",
        "tags": ["box", "parcel", "product", "inventory", "shipping", "delivery", "order", "goods"],
        "primitives": [
            rect(4, 5, 16, 15), line(4, 9, 20, 9), line(12, 5, 12, 20),
            line(4, 5, 12, 9), line(20, 5, 12, 9),
        ],
    },
    {
        "id": "warehouse",
        "name": "Warehouse",
        "category": "Operations",
        "tags": ["storage", "inventory", "distribution center", "logistics", "facility", "stock", "supply chain", "fulfillment"],
        "primitives": [
            *polyline((2, 9), (12, 3), (22, 9)), rect(4, 9, 16, 12), rect(8, 13, 8, 8),
            line(8, 16, 16, 16), line(8, 19, 16, 19),
        ],
    },
    {
        "id": "checklist",
        "name": "Checklist",
        "category": "Operations",
        "tags": ["tasks", "todo", "quality", "audit", "completion", "requirements", "control", "inspection"],
        "primitives": [
            rect(4, 3, 16, 18), *polyline((7, 8), (8.5, 9.5), (11, 6)), line(13, 8, 17, 8),
            *polyline((7, 13), (8.5, 14.5), (11, 11)), line(13, 13, 17, 13),
            *polyline((7, 18), (8.5, 19.5), (11, 16)), line(13, 18, 17, 18),
        ],
    },
    {
        "id": "process",
        "name": "Process",
        "category": "Operations",
        "tags": ["workflow", "steps", "flow", "procedure", "sequence", "operations", "pipeline", "value chain"],
        "primitives": [
            rect(2, 8, 5, 8), rect(9.5, 8, 5, 8), rect(17, 8, 5, 8),
            line(7, 12, 9.5, 12), line(8, 10.5, 9.5, 12), line(8, 13.5, 9.5, 12),
            line(14.5, 12, 17, 12), line(15.5, 10.5, 17, 12), line(15.5, 13.5, 17, 12),
        ],
    },
    {
        "id": "wrench",
        "name": "Maintenance",
        "category": "Operations",
        "tags": ["wrench", "repair", "service", "tool", "support", "fix", "engineering", "maintenance"],
        "primitives": [
            ellipse(3, 3, 7, 7), ellipse(14, 14, 7, 7), line(8, 8, 16, 16), line(6, 3, 6, 7), line(3, 6, 7, 6),
        ],
    },
    {
        "id": "supply-chain",
        "name": "Supply Chain",
        "category": "Operations",
        "tags": ["network", "supplier", "procurement", "logistics", "distribution", "value chain", "flow", "ecosystem"],
        "primitives": [
            rect(2, 3, 6, 5), rect(16, 3, 6, 5), rect(9, 16, 6, 5),
            line(8, 5.5, 16, 5.5), line(5, 8, 11, 16), line(19, 8, 13, 16),
        ],
    },
    {
        "id": "mail",
        "name": "Email",
        "category": "Communication",
        "tags": ["mail", "message", "inbox", "contact", "newsletter", "send", "correspondence", "communication"],
        "primitives": [
            rect(3, 5, 18, 14), line(3, 5, 12, 13), line(21, 5, 12, 13),
            line(3, 19, 9, 12), line(21, 19, 15, 12),
        ],
    },
    {
        "id": "chat",
        "name": "Conversation",
        "category": "Communication",
        "tags": ["chat", "message", "discussion", "feedback", "comment", "dialogue", "support", "communication"],
        "primitives": [
            rect(3, 4, 18, 13), *polyline((8, 17), (7, 21), (12, 17)),
            ellipse(7, 9.5, 1.5, 1.5, filled=True), ellipse(11.25, 9.5, 1.5, 1.5, filled=True),
            ellipse(15.5, 9.5, 1.5, 1.5, filled=True),
        ],
    },
    {
        "id": "phone",
        "name": "Call",
        "category": "Communication",
        "tags": ["phone", "telephone", "contact", "customer service", "voice", "support", "call center", "communication"],
        "primitives": [
            ellipse(4, 4, 5, 5), ellipse(15, 15, 5, 5),
            *polyline((7, 8), (9, 12), (12, 15), (16, 17)), line(4, 7, 7, 13), line(17, 11, 20, 17),
        ],
    },
    {
        "id": "megaphone",
        "name": "Announcement",
        "category": "Communication",
        "tags": ["megaphone", "marketing", "broadcast", "campaign", "promotion", "news", "advertising", "communication"],
        "primitives": [
            *polyline((3, 10), (8, 10), (20, 4), (20, 18), (8, 14), (3, 14), (3, 10), closed=True),
            line(8, 14, 10, 21), line(10, 21, 14, 21), line(14, 21, 12, 15),
        ],
    },
]

ALIASES_BY_ID: dict[str, list[str]] = {
    "analytics": ["dashboard", "business intelligence", "bi"],
    "briefcase": ["work", "career", "portfolio"],
    "target": ["bullseye", "aim", "north star"],
    "users": ["people", "group", "workforce"],
    "cloud": ["cloud computing", "public cloud", "private cloud"],
    "database": ["datastore", "data store", "repository"],
    "server": ["data center", "compute instance", "host"],
    "laptop": ["notebook", "computer", "pc"],
    "mobile": ["smartphone", "cell phone", "handset"],
    "chip": ["processor", "cpu", "gpu"],
    "network": ["topology", "mesh", "graph"],
    "lock": ["cybersecurity", "secure", "access control"],
    "presentation": ["slide deck", "slideshow", "pitch deck"],
    "calendar": ["planner", "agenda", "schedule"],
    "clock": ["time", "timer", "deadline"],
    "lightbulb": ["idea", "innovation", "insight"],
    "compass": ["direction", "navigation", "strategy"],
    "flag": ["milestone", "finish", "priority"],
    "puzzle": ["solution", "integration", "fit"],
    "handshake": ["partnership", "agreement", "deal"],
    "currency": ["money", "dollar", "usd"],
    "percent": ["percentage", "ratio", "rate"],
    "bank": ["financial institution", "treasury", "capital"],
    "wallet": ["payment", "budget", "expense"],
    "calculator": ["calculation", "math", "accounting"],
    "invoice": ["bill", "receipt", "statement"],
    "trend-up": ["growth", "increase", "upward trend"],
    "trend-down": ["decline", "decrease", "downward trend"],
    "code": ["source code", "programming", "development"],
    "terminal": ["command line", "console", "shell"],
    "api": ["application programming interface", "endpoint", "service"],
    "automation": ["workflow automation", "rpa", "orchestration"],
    "globe": ["world", "international", "global"],
    "wifi": ["wireless", "signal", "connectivity"],
    "layers": ["stack", "architecture layers", "technology stack"],
    "webhook": ["callback", "event trigger", "integration hook"],
    "factory": ["manufacturing plant", "production", "industrial"],
    "truck": ["delivery", "freight", "transport"],
    "package": ["parcel", "box", "shipment"],
    "warehouse": ["distribution center", "storage", "fulfillment"],
    "checklist": ["task list", "todo", "audit"],
    "process": ["workflow", "procedure", "flow"],
    "wrench": ["maintenance", "repair", "service"],
    "supply-chain": ["value chain", "logistics network", "procurement"],
    "mail": ["email", "inbox", "message"],
    "chat": ["conversation", "discussion", "dialogue"],
    "phone": ["call", "telephone", "voice"],
    "megaphone": ["announcement", "broadcast", "campaign"],
}

EXTRA_TAGS_BY_ID: dict[str, list[str]] = {
    "analytics": ["dashboard", "business intelligence", "kpi"],
    "briefcase": ["corporate", "career", "professional"],
    "target": ["aim", "bullseye", "success"],
    "users": ["employee", "workforce", "group"],
    "cloud": ["cloud computing", "public cloud", "private cloud"],
    "database": ["datastore", "records", "repository"],
    "server": ["data center", "hardware", "virtual machine"],
    "laptop": ["notebook", "pc", "remote work"],
    "mobile": ["cellular", "iphone", "android"],
    "chip": ["neural processing unit", "gpu", "hardware"],
    "network": ["topology", "graph", "mesh"],
    "lock": ["secure", "access control", "encryption"],
}


def transformed(primitives: list[dict[str, Any]], scale: float, dx: float, dy: float) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for primitive in primitives:
        item = dict(primitive)
        if primitive["kind"] == "line":
            for axis in ("x1", "x2"):
                item[axis] = primitive[axis] * scale + dx
            for axis in ("y1", "y2"):
                item[axis] = primitive[axis] * scale + dy
        else:
            item["x"] = primitive["x"] * scale + dx
            item["y"] = primitive["y"] * scale + dy
            item["width"] = primitive["width"] * scale
            item["height"] = primitive["height"] * scale
        result.append(item)
    return result


BADGES: dict[str, list[dict[str, Any]]] = {
    "add": [ellipse(16, 16, 7, 7), line(19.5, 17.5, 19.5, 21.5), line(17.5, 19.5, 21.5, 19.5)],
    "approved": [ellipse(16, 16, 7, 7), *polyline((17.5, 19.5), (19, 21), (21.5, 17.5))],
    "remove": [ellipse(16, 16, 7, 7), line(17.5, 19.5, 21.5, 19.5)],
    "blocked": [ellipse(16, 16, 7, 7), line(17.5, 17.5, 21.5, 21.5), line(21.5, 17.5, 17.5, 21.5)],
    "alert": [ellipse(16, 16, 7, 7), line(19.5, 17.5, 19.5, 20), ellipse(18.75, 21, 1.5, 1.5, filled=True)],
    "search": [ellipse(16, 16, 5.5, 5.5), line(20.25, 20.25, 23, 23)],
    "secure": [ellipse(17.25, 14.5, 4.5, 6), rect(16, 18, 7, 5), ellipse(18.75, 19.5, 1.5, 1.5, filled=True)],
    "growth": [ellipse(16, 16, 7, 7), line(19.5, 21.5, 19.5, 17.5), line(17.5, 19.5, 19.5, 17.5), line(21.5, 19.5, 19.5, 17.5)],
}

VARIANTS = [
    ("add", "Add", ["new", "create", "plus", "onboard"]),
    ("approved", "Approved", ["approved", "complete", "verified", "check"]),
    ("remove", "Remove", ["remove", "subtract", "reduce", "minus"]),
    ("blocked", "Blocked", ["blocked", "failed", "cancel", "unavailable"]),
    ("alert", "Alert", ["alert", "warning", "issue", "attention"]),
    ("search", "Search", ["search", "find", "discover", "inspect"]),
    ("secure", "Secure", ["secure", "protected", "private", "controlled"]),
    ("growth", "Growth", ["growth", "increase", "scale", "improve"]),
]

NEW_BASE_ICONS = [
    {
        "id": "document",
        "name": "Document",
        "category": "Business",
        "aliases": ["file", "paper", "page"],
        "tags": ["document", "file", "paper", "report", "memo", "content", "record", "attachment"],
        "primitives": [rect(5, 2, 14, 20), *polyline((14, 2), (14, 7), (19, 7)), line(8, 11, 16, 11), line(8, 15, 16, 15), line(8, 19, 13, 19)],
    },
    {
        "id": "folder",
        "name": "Folder",
        "category": "Business",
        "aliases": ["directory", "files", "archive"],
        "tags": ["folder", "directory", "files", "archive", "storage", "collection", "workspace", "repository"],
        "primitives": [*polyline((3, 7), (9, 7), (11, 10), (21, 10), (20, 20), (4, 20), (3, 7), closed=True), line(3, 10, 21, 10)],
    },
    {
        "id": "user",
        "name": "Person",
        "category": "People",
        "aliases": ["user", "employee", "individual"],
        "tags": ["person", "user", "employee", "customer", "individual", "profile", "stakeholder", "human"],
        "primitives": [ellipse(8, 3, 8, 8), *polyline((4, 21), (5, 16), (9, 13), (15, 13), (19, 16), (20, 21))],
    },
    {
        "id": "shield",
        "name": "Shield",
        "category": "Security",
        "aliases": ["protection", "defense", "safety"],
        "tags": ["shield", "protection", "security", "defense", "safety", "risk", "control", "resilience"],
        "primitives": [*polyline((12, 2), (20, 5), (19, 14), (16, 19), (12, 22), (8, 19), (5, 14), (4, 5), (12, 2), closed=True)],
    },
    {
        "id": "building",
        "name": "Office",
        "category": "Business",
        "aliases": ["building", "headquarters", "company"],
        "tags": ["office", "building", "headquarters", "company", "corporate", "organization", "facility", "workplace"],
        "primitives": [
            rect(5, 3, 14, 18), rect(8, 6, 2, 2, filled=True), rect(14, 6, 2, 2, filled=True),
            rect(8, 11, 2, 2, filled=True), rect(14, 11, 2, 2, filled=True), rect(10, 16, 4, 5),
        ],
    },
    {
        "id": "sustainability",
        "name": "Sustainability",
        "category": "ESG",
        "aliases": ["leaf", "green", "environment"],
        "tags": ["sustainability", "leaf", "green", "environment", "esg", "nature", "climate", "ecology"],
        "primitives": [
            *polyline((4, 18), (5, 10), (10, 5), (20, 3), (19, 13), (14, 18), (4, 18), closed=True),
            line(5, 18, 17, 6), line(10, 13, 15, 13), line(13, 10, 13, 7),
        ],
    },
    {
        "id": "energy",
        "name": "Energy",
        "category": "ESG",
        "aliases": ["power", "electricity", "lightning"],
        "tags": ["energy", "power", "electricity", "lightning", "utility", "renewable", "capacity", "grid"],
        "primitives": [*polyline((13, 2), (6, 13), (11, 13), (9, 22), (18, 10), (13, 10), (13, 2), closed=True)],
    },
    {
        "id": "risk",
        "name": "Risk",
        "category": "Security",
        "aliases": ["warning", "hazard", "issue"],
        "tags": ["risk", "warning", "hazard", "issue", "danger", "exposure", "threat", "attention"],
        "primitives": [
            *polyline((12, 2), (22, 21), (2, 21), (12, 2), closed=True),
            line(12, 8, 12, 15), ellipse(11, 17.5, 2, 2, filled=True),
        ],
    },
    {
        "id": "org-chart",
        "name": "Organization Chart",
        "category": "People",
        "aliases": ["org chart", "hierarchy", "reporting lines"],
        "tags": ["organization", "org chart", "hierarchy", "structure", "reporting", "team", "management", "operating model"],
        "primitives": [
            rect(9, 2, 6, 5), rect(2, 17, 5, 5), rect(9.5, 17, 5, 5), rect(17, 17, 5, 5),
            line(12, 7, 12, 12), line(4.5, 12, 19.5, 12), line(4.5, 12, 4.5, 17),
            line(12, 12, 12, 17), line(19.5, 12, 19.5, 17),
        ],
    },
    {
        "id": "funnel",
        "name": "Funnel",
        "category": "Business",
        "aliases": ["sales funnel", "conversion funnel", "filter"],
        "tags": ["funnel", "sales", "marketing", "conversion", "pipeline", "filter", "customer journey", "lead"],
        "primitives": [*polyline((3, 4), (21, 4), (15, 12), (15, 20), (9, 17), (9, 12), (3, 4), closed=True), line(6, 8, 18, 8)],
    },
    {
        "id": "rocket",
        "name": "Launch",
        "category": "Business",
        "aliases": ["rocket", "go to market", "takeoff"],
        "tags": ["launch", "rocket", "go to market", "takeoff", "startup", "growth", "initiative", "acceleration"],
        "primitives": [
            ellipse(8, 2, 8, 16), line(8, 13, 4, 18), line(16, 13, 20, 18),
            *polyline((10, 18), (12, 23), (14, 18)), ellipse(10.5, 6, 3, 3),
        ],
    },
    {
        "id": "trophy",
        "name": "Achievement",
        "category": "Business",
        "aliases": ["trophy", "winner", "award"],
        "tags": ["achievement", "trophy", "winner", "award", "success", "recognition", "best", "performance"],
        "primitives": [
            ellipse(6, 3, 12, 10), line(4, 5, 6, 9), line(20, 5, 18, 9),
            line(12, 13, 12, 18), line(8, 18, 16, 18), line(7, 21, 17, 21),
        ],
    },
    {
        "id": "credit-card",
        "name": "Credit Card",
        "category": "Finance",
        "aliases": ["payment card", "debit card", "card payment"],
        "tags": ["credit card", "payment", "debit", "finance", "transaction", "purchase", "billing", "banking"],
        "primitives": [rect(2, 5, 20, 14), line(2, 9, 22, 9), rect(5, 13, 5, 2, filled=True), line(13, 15, 19, 15)],
    },
    {
        "id": "pie-chart",
        "name": "Pie Chart",
        "category": "Finance",
        "aliases": ["market share", "portfolio mix", "composition"],
        "tags": ["pie chart", "share", "mix", "composition", "portfolio", "market share", "allocation", "percentage"],
        "primitives": [ellipse(3, 3, 18, 18), line(12, 12, 12, 3), line(12, 12, 20, 8), line(12, 12, 18, 19)],
    },
    {
        "id": "balance-scale",
        "name": "Balance",
        "category": "Finance",
        "aliases": ["scales", "tradeoff", "justice"],
        "tags": ["balance", "scales", "tradeoff", "justice", "decision", "comparison", "equilibrium", "governance"],
        "primitives": [
            line(12, 3, 12, 20), line(5, 7, 19, 7), line(8, 21, 16, 21),
            line(6, 7, 3, 14), line(6, 7, 9, 14), line(18, 7, 15, 14), line(18, 7, 21, 14),
            line(3, 14, 9, 14), line(15, 14, 21, 14),
        ],
    },
    {
        "id": "location",
        "name": "Location",
        "category": "Operations",
        "aliases": ["map pin", "place", "site"],
        "tags": ["location", "map pin", "place", "site", "geography", "office", "market", "destination"],
        "primitives": [
            ellipse(5, 2, 14, 14), *polyline((6, 12), (12, 22), (18, 12)),
            ellipse(9.5, 6.5, 5, 5),
        ],
    },
    {
        "id": "route",
        "name": "Route",
        "category": "Operations",
        "aliases": ["path", "journey", "roadmap"],
        "tags": ["route", "path", "journey", "roadmap", "logistics", "direction", "travel", "sequence"],
        "primitives": [
            ellipse(2, 16, 5, 5), ellipse(17, 2, 5, 5),
            *polyline((6, 18), (11, 18), (14, 14), (10, 10), (13, 6), (18, 5)),
        ],
    },
    {
        "id": "robot",
        "name": "Robot",
        "category": "Technology",
        "aliases": ["bot", "robotics", "automation bot"],
        "tags": ["robot", "bot", "robotics", "automation", "technology", "machine", "assistant", "autonomous"],
        "primitives": [
            rect(4, 6, 16, 12), line(12, 3, 12, 6), ellipse(10.5, 1.5, 3, 3, filled=True),
            ellipse(7, 10, 2, 2, filled=True), ellipse(15, 10, 2, 2, filled=True),
            line(8, 15, 16, 15), line(1, 10, 4, 10), line(20, 10, 23, 10),
        ],
    },
    {
        "id": "ai-brain",
        "name": "AI Brain",
        "category": "Technology",
        "aliases": ["artificial intelligence", "neural network", "machine intelligence"],
        "tags": ["ai", "artificial intelligence", "brain", "neural network", "machine learning", "cognition", "model", "intelligence"],
        "primitives": [
            ellipse(3, 7, 7, 8), ellipse(7, 3, 7, 9), ellipse(11, 3, 7, 9), ellipse(15, 7, 6, 8),
            ellipse(5, 12, 7, 8), ellipse(12, 12, 7, 8), line(12, 4, 12, 20),
        ],
    },
    {
        "id": "data-pipeline",
        "name": "Data Pipeline",
        "category": "Technology",
        "aliases": ["etl", "data flow", "integration pipeline"],
        "tags": ["data pipeline", "etl", "data flow", "integration", "processing", "analytics", "transform", "data engineering"],
        "primitives": [
            ellipse(2, 8, 6, 4), ellipse(16, 3, 6, 4), ellipse(16, 17, 6, 4),
            line(8, 10, 16, 5), line(8, 10, 16, 19), line(13, 6, 16, 5), line(14, 3, 16, 5),
            line(13, 18, 16, 19), line(14, 21, 16, 19),
        ],
    },
    {
        "id": "container",
        "name": "Container",
        "category": "Technology",
        "aliases": ["docker", "software container", "shipping container"],
        "tags": ["container", "docker", "software", "deployment", "cloud native", "microservice", "platform", "infrastructure"],
        "primitives": [
            rect(2, 5, 20, 14), line(6, 5, 6, 19), line(10, 5, 10, 19),
            line(14, 5, 14, 19), line(18, 5, 18, 19),
        ],
    },
    {
        "id": "firewall",
        "name": "Firewall",
        "category": "Security",
        "aliases": ["network security", "security wall", "perimeter"],
        "tags": ["firewall", "network security", "perimeter", "cyber security", "protection", "access", "control", "defense"],
        "primitives": [
            rect(2, 4, 20, 16), line(2, 9, 22, 9), line(2, 14, 22, 14),
            line(7, 4, 7, 9), line(16, 4, 16, 9), line(11, 9, 11, 14),
            line(7, 14, 7, 20), line(16, 14, 16, 20),
        ],
    },
    {
        "id": "key",
        "name": "Key",
        "category": "Security",
        "aliases": ["access key", "credentials", "permission"],
        "tags": ["key", "access", "credentials", "permission", "authentication", "security", "password", "identity"],
        "primitives": [
            ellipse(3, 3, 9, 9), line(10, 10, 21, 21), line(16, 16, 19, 13), line(19, 19, 22, 16),
        ],
    },
    {
        "id": "certificate",
        "name": "Certificate",
        "category": "Security",
        "aliases": ["compliance certificate", "accreditation", "verified document"],
        "tags": ["certificate", "compliance", "accreditation", "verified", "audit", "qualification", "standard", "governance"],
        "primitives": [
            rect(4, 2, 16, 16), line(8, 6, 16, 6), line(8, 10, 14, 10),
            ellipse(9, 14, 6, 6), line(10, 19, 9, 23), line(14, 19, 15, 23),
        ],
    },
    {
        "id": "video",
        "name": "Video",
        "category": "Communication",
        "aliases": ["video call", "camera", "recording"],
        "tags": ["video", "camera", "recording", "meeting", "conference", "media", "webinar", "communication"],
        "primitives": [rect(2, 6, 14, 12), *polyline((16, 10), (22, 7), (22, 17), (16, 14), closed=True), line(8, 9, 8, 15)],
    },
    {
        "id": "microphone",
        "name": "Microphone",
        "category": "Communication",
        "aliases": ["mic", "voice", "speech"],
        "tags": ["microphone", "voice", "speech", "audio", "podcast", "presentation", "speaker", "recording"],
        "primitives": [
            ellipse(8, 2, 8, 13), *polyline((5, 10), (5, 14), (8, 18), (12, 19), (16, 18), (19, 14), (19, 10)),
            line(12, 19, 12, 22), line(8, 22, 16, 22),
        ],
    },
    {
        "id": "bell",
        "name": "Notification",
        "category": "Communication",
        "aliases": ["bell", "reminder", "alert"],
        "tags": ["notification", "bell", "reminder", "alert", "update", "attention", "message", "subscription"],
        "primitives": [
            ellipse(6, 4, 12, 14), line(4, 18, 20, 18), line(7, 18, 5, 15), line(17, 18, 19, 15),
            ellipse(10.5, 19, 3, 3),
        ],
    },
    {
        "id": "send",
        "name": "Send",
        "category": "Communication",
        "aliases": ["paper plane", "submit", "forward"],
        "tags": ["send", "paper plane", "submit", "forward", "message", "share", "email", "delivery"],
        "primitives": [
            *polyline((2, 11), (22, 3), (16, 21), (11, 14), (2, 11), closed=True),
            line(11, 14, 22, 3), line(9, 14, 9, 19), line(9, 19, 13, 16),
        ],
    },
    {
        "id": "solar",
        "name": "Solar Energy",
        "category": "ESG",
        "aliases": ["solar panel", "renewable power", "photovoltaic"],
        "tags": ["solar", "solar panel", "renewable", "energy", "photovoltaic", "electricity", "decarbonization", "climate"],
        "primitives": [
            rect(3, 10, 14, 10), line(3, 15, 17, 15), line(8, 10, 8, 20), line(13, 10, 13, 20),
            ellipse(17, 2, 5, 5), line(19.5, 0, 19.5, 2), line(16, 5, 14, 7),
        ],
    },
    {
        "id": "wind",
        "name": "Wind Energy",
        "category": "ESG",
        "aliases": ["wind turbine", "wind power", "renewable energy"],
        "tags": ["wind", "wind turbine", "renewable", "energy", "power", "electricity", "climate", "decarbonization"],
        "primitives": [
            ellipse(10.5, 8.5, 3, 3, filled=True), line(12, 12, 12, 22),
            line(12, 10, 12, 2), line(12, 10, 4, 14), line(12, 10, 20, 14),
            line(12, 2, 15, 6), line(4, 14, 9, 14), line(20, 14, 15, 11),
        ],
    },
    {
        "id": "water",
        "name": "Water",
        "category": "ESG",
        "aliases": ["water drop", "water resource", "liquid"],
        "tags": ["water", "drop", "resource", "liquid", "environment", "sustainability", "utility", "conservation"],
        "primitives": [
            *polyline((12, 2), (5, 12), (5, 16), (8, 20), (12, 22), (16, 20), (19, 16), (19, 12), (12, 2), closed=True),
            line(9, 17, 11, 19), line(11, 19, 14, 18),
        ],
    },
    {
        "id": "recycle",
        "name": "Circular Economy",
        "category": "ESG",
        "aliases": ["recycle", "circularity", "reuse"],
        "tags": ["circular economy", "recycle", "circularity", "reuse", "waste", "sustainability", "materials", "resource efficiency"],
        "primitives": [
            *polyline((12, 3), (17, 7), (14, 7)), *polyline((18, 8), (21, 14), (18, 14)),
            *polyline((17, 17), (11, 21), (11, 18)), *polyline((9, 20), (3, 16), (6, 15)),
            *polyline((4, 13), (7, 6), (9, 9)), line(9, 6, 15, 6), line(19, 14, 16, 19), line(11, 19, 5, 15),
        ],
    },
    {
        "id": "battery",
        "name": "Battery",
        "category": "ESG",
        "aliases": ["energy storage", "power storage", "charge"],
        "tags": ["battery", "energy storage", "charge", "power", "electricity", "capacity", "renewable", "grid"],
        "primitives": [
            rect(3, 6, 17, 12), rect(20, 10, 2, 4), rect(6, 9, 3, 6, filled=True),
            rect(11, 9, 3, 6, filled=True), rect(16, 9, 2, 6, filled=True),
        ],
    },
    {
        "id": "clipboard",
        "name": "Clipboard",
        "category": "Business",
        "aliases": ["notes", "task board", "copy"],
        "tags": ["clipboard", "notes", "tasks", "copy", "record", "checklist", "document", "workflow"],
        "primitives": [rect(4, 4, 16, 18), rect(8, 2, 8, 5), line(8, 10, 16, 10), line(8, 14, 16, 14), line(8, 18, 13, 18)],
    },
    {
        "id": "book",
        "name": "Knowledge",
        "category": "Business",
        "aliases": ["book", "knowledge base", "manual"],
        "tags": ["book", "knowledge", "manual", "learning", "playbook", "documentation", "training", "reference"],
        "primitives": [rect(2, 4, 10, 17), rect(12, 4, 10, 17), line(12, 4, 12, 21), line(5, 8, 9, 8), line(15, 8, 19, 8)],
    },
    {
        "id": "link",
        "name": "Link",
        "category": "Technology",
        "aliases": ["connection", "hyperlink", "chain"],
        "tags": ["link", "connection", "hyperlink", "chain", "integration", "relationship", "url", "dependency"],
        "primitives": [
            ellipse(2, 7, 10, 8), ellipse(12, 9, 10, 8), line(8, 12, 16, 12), line(5, 10, 3, 12), line(19, 14, 21, 12),
        ],
    },
    {
        "id": "upload",
        "name": "Upload",
        "category": "Technology",
        "aliases": ["import", "send file", "move up"],
        "tags": ["upload", "import", "send", "file", "cloud", "transfer", "data", "arrow up"],
        "primitives": [line(12, 4, 12, 17), line(7, 9, 12, 4), line(17, 9, 12, 4), *polyline((4, 15), (4, 21), (20, 21), (20, 15))],
    },
    {
        "id": "download",
        "name": "Download",
        "category": "Technology",
        "aliases": ["export", "receive file", "move down"],
        "tags": ["download", "export", "receive", "file", "cloud", "transfer", "data", "arrow down"],
        "primitives": [line(12, 3, 12, 16), line(7, 11, 12, 16), line(17, 11, 12, 16), *polyline((4, 15), (4, 21), (20, 21), (20, 15))],
    },
    {
        "id": "table",
        "name": "Table",
        "category": "Business",
        "aliases": ["spreadsheet", "data table", "grid"],
        "tags": ["table", "spreadsheet", "grid", "data", "matrix", "rows", "columns", "analysis"],
        "primitives": [
            rect(2, 4, 20, 16), line(2, 9, 22, 9), line(2, 14, 22, 14),
            line(8, 4, 8, 20), line(15, 4, 15, 20),
        ],
    },
]

ICONS.extend(NEW_BASE_ICONS)

CONSULTING_ICONS = [
    {
        "id": "roadmap",
        "name": "Roadmap",
        "category": "Business",
        "aliases": ["strategic roadmap", "initiative plan", "delivery roadmap"],
        "tags": ["roadmap", "strategy", "plan", "timeline", "milestone", "initiative", "delivery", "transformation"],
        "primitives": [line(3, 18, 21, 18), line(5, 18, 5, 8), line(12, 18, 12, 4), line(19, 18, 19, 10), ellipse(3, 5, 4, 4, filled=True), ellipse(10, 2, 4, 4, filled=True), ellipse(17, 7, 4, 4, filled=True)],
    },
    {
        "id": "strategy-map",
        "name": "Strategy Map",
        "category": "Business",
        "aliases": ["strategic plan", "strategy cascade", "objective map"],
        "tags": ["strategy map", "strategy", "objectives", "cascade", "alignment", "priorities", "execution", "balanced scorecard"],
        "primitives": [rect(2, 3, 8, 5), rect(14, 3, 8, 5), rect(8, 16, 8, 5), line(6, 8, 10, 16), line(18, 8, 14, 16), line(10, 5.5, 14, 5.5)],
    },
    {
        "id": "portfolio",
        "name": "Portfolio",
        "category": "Business",
        "aliases": ["initiative portfolio", "project portfolio", "business portfolio"],
        "tags": ["portfolio", "initiatives", "projects", "programs", "prioritization", "investment", "pipeline", "governance"],
        "primitives": [rect(3, 5, 8, 6), rect(13, 5, 8, 6), rect(3, 13, 8, 6), rect(13, 13, 8, 6), line(7, 3, 17, 3), line(7, 3, 7, 5), line(17, 3, 17, 5)],
    },
    {
        "id": "matrix",
        "name": "Two by Two Matrix",
        "category": "Business",
        "aliases": ["2x2 matrix", "quadrant chart", "prioritization matrix"],
        "tags": ["matrix", "2x2", "quadrant", "prioritization", "framework", "analysis", "positioning", "decision"],
        "primitives": [rect(3, 3, 18, 18), line(12, 3, 12, 21), line(3, 12, 21, 12), ellipse(16, 6, 2, 2, filled=True), ellipse(7, 16, 2, 2, filled=True)],
    },
    {
        "id": "swot",
        "name": "SWOT Analysis",
        "category": "Business",
        "aliases": ["swot matrix", "strengths weaknesses", "strategic assessment"],
        "tags": ["swot", "strengths", "weaknesses", "opportunities", "threats", "strategy", "assessment", "framework"],
        "primitives": [rect(2, 2, 20, 20), line(12, 2, 12, 22), line(2, 12, 22, 12), line(5, 7, 9, 7), line(15, 7, 19, 7), line(5, 17, 9, 17), line(15, 17, 19, 17)],
    },
    {
        "id": "decision-tree",
        "name": "Decision Tree",
        "category": "Business",
        "aliases": ["decision logic", "choice tree", "branching analysis"],
        "tags": ["decision tree", "decision", "choice", "options", "branch", "logic", "scenario", "analysis"],
        "primitives": [ellipse(2, 8, 6, 6), rect(16, 2, 6, 5), rect(16, 10, 6, 5), rect(16, 18, 6, 4), line(8, 11, 12, 11), line(12, 4.5, 12, 20), line(12, 4.5, 16, 4.5), line(12, 12.5, 16, 12.5), line(12, 20, 16, 20)],
    },
    {
        "id": "milestone",
        "name": "Milestone",
        "category": "Business",
        "aliases": ["project milestone", "stage gate", "checkpoint"],
        "tags": ["milestone", "checkpoint", "stage gate", "project", "timeline", "delivery", "progress", "deadline"],
        "primitives": [line(3, 17, 21, 17), ellipse(4, 14, 6, 6, filled=True), *polyline((12, 17), (16, 12), (20, 17), (16, 22), (12, 17), closed=True), line(7, 6, 7, 14), *polyline((7, 6), (14, 6), (11, 9), (7, 9))],
    },
    {
        "id": "speedometer",
        "name": "Performance Gauge",
        "category": "Business",
        "aliases": ["speedometer", "kpi gauge", "performance meter"],
        "tags": ["performance", "gauge", "speedometer", "kpi", "score", "measurement", "dashboard", "progress"],
        "primitives": [ellipse(3, 5, 18, 18), line(4, 17, 20, 17), line(12, 17, 17, 9), ellipse(10.5, 15.5, 3, 3, filled=True), line(6, 13, 8, 14), line(12, 8, 12, 10), line(18, 13, 16, 14)],
    },
    {
        "id": "binoculars",
        "name": "Market Outlook",
        "category": "Business",
        "aliases": ["binoculars", "market scan", "future outlook"],
        "tags": ["outlook", "binoculars", "market scan", "future", "vision", "research", "opportunity", "horizon"],
        "primitives": [ellipse(2, 11, 9, 9), ellipse(13, 11, 9, 9), rect(5, 5, 5, 9), rect(14, 5, 5, 9), line(10, 8, 14, 8), line(10, 12, 14, 12)],
    },
    {
        "id": "mountain",
        "name": "Ambition",
        "category": "Business",
        "aliases": ["mountain", "aspiration", "summit"],
        "tags": ["ambition", "mountain", "aspiration", "summit", "challenge", "goal", "vision", "achievement"],
        "primitives": [*polyline((2, 21), (9, 8), (13, 14), (17, 5), (23, 21)), line(2, 21, 23, 21), *polyline((14, 10), (17, 5), (20, 10)), line(8, 10, 10, 12)],
    },
    {
        "id": "diamond",
        "name": "Value",
        "category": "Business",
        "aliases": ["diamond", "premium value", "value proposition"],
        "tags": ["value", "diamond", "premium", "proposition", "benefit", "differentiation", "quality", "customer"],
        "primitives": [*polyline((3, 8), (7, 3), (17, 3), (21, 8), (12, 21), (3, 8), closed=True), line(3, 8, 21, 8), line(7, 3, 9, 8), line(17, 3, 15, 8), line(9, 8, 12, 21), line(15, 8, 12, 21)],
    },
    {
        "id": "star",
        "name": "Priority",
        "category": "Business",
        "aliases": ["star", "favorite", "top priority"],
        "tags": ["priority", "star", "favorite", "important", "rating", "highlight", "top", "focus"],
        "primitives": [*polyline((12, 2), (15, 9), (22, 9), (16.5, 13.5), (19, 21), (12, 16.5), (5, 21), (7.5, 13.5), (2, 9), (9, 9), (12, 2), closed=True)],
    },
    {
        "id": "question",
        "name": "Key Question",
        "category": "Business",
        "aliases": ["question mark", "business question", "unknown"],
        "tags": ["question", "question mark", "unknown", "issue", "hypothesis", "inquiry", "problem", "clarification"],
        "primitives": [ellipse(3, 3, 18, 18), *polyline((8, 8), (9, 6), (12, 5), (15, 6), (16, 9), (14, 12), (12, 13), (12, 16)), ellipse(11, 18, 2, 2, filled=True)],
    },
    {
        "id": "bookmark",
        "name": "Bookmark",
        "category": "Business",
        "aliases": ["saved item", "reference marker", "favorite page"],
        "tags": ["bookmark", "save", "reference", "favorite", "marker", "page", "knowledge", "content"],
        "primitives": [*polyline((6, 3), (18, 3), (18, 21), (12, 16), (6, 21), (6, 3), closed=True), line(9, 7, 15, 7)],
    },
    {
        "id": "scenario",
        "name": "Scenario Planning",
        "category": "Business",
        "aliases": ["scenario analysis", "future paths", "strategic options"],
        "tags": ["scenario", "planning", "future", "options", "uncertainty", "strategy", "paths", "forecast"],
        "primitives": [line(3, 12, 9, 12), line(9, 12, 15, 5), line(9, 12, 15, 12), line(9, 12, 15, 19), line(15, 5, 21, 5), line(15, 12, 21, 12), line(15, 19, 21, 19), ellipse(1, 10, 4, 4, filled=True)],
    },
    {
        "id": "coin-stack",
        "name": "Capital",
        "category": "Finance",
        "aliases": ["coin stack", "cash reserves", "financial capital"],
        "tags": ["capital", "coins", "cash", "reserves", "funding", "money", "liquidity", "finance"],
        "primitives": [ellipse(3, 15, 11, 5), ellipse(3, 11, 11, 5), ellipse(3, 7, 11, 5), ellipse(10, 15, 11, 5), ellipse(10, 11, 11, 5), line(3, 9.5, 3, 17.5), line(14, 9.5, 14, 17.5), line(21, 13.5, 21, 17.5)],
    },
    {
        "id": "cash-flow",
        "name": "Cash Flow",
        "category": "Finance",
        "aliases": ["money flow", "funds flow", "cash movement"],
        "tags": ["cash flow", "money", "inflow", "outflow", "liquidity", "finance", "working capital", "treasury"],
        "primitives": [ellipse(8, 7, 8, 8), line(12, 9, 12, 13), line(10, 11, 14, 11), line(2, 6, 8, 6), line(5, 3, 8, 6), line(5, 9, 8, 6), line(16, 18, 22, 18), line(19, 15, 22, 18), line(19, 21, 22, 18)],
    },
    {
        "id": "budget",
        "name": "Budget",
        "category": "Finance",
        "aliases": ["budget plan", "spending plan", "financial plan"],
        "tags": ["budget", "plan", "spending", "cost", "finance", "allocation", "forecast", "control"],
        "primitives": [rect(4, 3, 16, 18), line(7, 8, 17, 8), line(7, 12, 17, 12), line(7, 16, 12, 16), ellipse(14, 14, 4, 4), line(16, 15, 16, 17), line(15, 16, 17, 16)],
    },
    {
        "id": "forecast",
        "name": "Financial Forecast",
        "category": "Finance",
        "aliases": ["forecast model", "financial projection", "outlook model"],
        "tags": ["forecast", "projection", "outlook", "finance", "model", "estimate", "planning", "future"],
        "primitives": [line(3, 20, 21, 20), line(3, 20, 3, 4), *polyline((5, 16), (9, 13), (12, 15), (16, 9)), line(16, 9, 21, 5), line(18, 5, 21, 5), line(21, 5, 21, 8), line(16, 9, 16, 20)],
    },
    {
        "id": "profit-loss",
        "name": "Profit and Loss",
        "category": "Finance",
        "aliases": ["p and l", "income statement", "profit loss statement"],
        "tags": ["profit", "loss", "income statement", "p&l", "revenue", "expense", "earnings", "finance"],
        "primitives": [rect(3, 3, 18, 18), line(7, 8, 17, 8), line(7, 12, 17, 12), line(7, 16, 13, 16), line(16, 14, 16, 18), line(14, 16, 18, 16), line(8, 6, 8, 10)],
    },
    {
        "id": "investment",
        "name": "Investment",
        "category": "Finance",
        "aliases": ["capital investment", "investing", "growth capital"],
        "tags": ["investment", "capital", "funding", "return", "growth", "portfolio", "finance", "asset"],
        "primitives": [ellipse(3, 12, 9, 9), line(7.5, 14, 7.5, 19), line(5.5, 16.5, 9.5, 16.5), *polyline((11, 15), (15, 11), (18, 13), (22, 6)), line(18, 6, 22, 6), line(22, 6, 22, 10)],
    },
    {
        "id": "tax",
        "name": "Tax",
        "category": "Finance",
        "aliases": ["taxation", "tax rate", "fiscal charge"],
        "tags": ["tax", "taxation", "fiscal", "rate", "government", "finance", "compliance", "payment"],
        "primitives": [rect(4, 3, 16, 18), ellipse(7, 7, 3, 3), ellipse(14, 14, 3, 3), line(8, 17, 16, 7), line(7, 12, 17, 12)],
    },
    {
        "id": "safe",
        "name": "Financial Security",
        "category": "Finance",
        "aliases": ["safe", "vault", "secure funds"],
        "tags": ["safe", "vault", "security", "money", "assets", "treasury", "protection", "finance"],
        "primitives": [rect(3, 3, 18, 18), rect(6, 6, 12, 12), ellipse(8, 8, 8, 8), line(12, 8, 12, 12), line(12, 12, 15, 14), rect(1, 8, 2, 8), rect(21, 8, 2, 8)],
    },
    {
        "id": "saas",
        "name": "Software as a Service",
        "category": "Technology",
        "aliases": ["saas", "cloud software", "subscription software"],
        "tags": ["saas", "software as a service", "cloud", "subscription", "application", "platform", "digital", "technology"],
        "primitives": [ellipse(3, 8, 7, 7), ellipse(7, 4, 10, 10), ellipse(14, 8, 7, 7), line(6, 15, 18, 15), rect(8, 12, 8, 8), line(10, 16, 14, 16)],
    },
    {
        "id": "microservices",
        "name": "Microservices",
        "category": "Technology",
        "aliases": ["service architecture", "distributed services", "microservice mesh"],
        "tags": ["microservices", "services", "architecture", "distributed", "api", "cloud native", "components", "software"],
        "primitives": [rect(2, 3, 6, 5), rect(16, 3, 6, 5), rect(2, 16, 6, 5), rect(16, 16, 6, 5), rect(9, 9, 6, 6), line(8, 5.5, 12, 9), line(16, 5.5, 12, 9), line(8, 18.5, 12, 15), line(16, 18.5, 12, 15)],
    },
    {
        "id": "blockchain",
        "name": "Blockchain",
        "category": "Technology",
        "aliases": ["distributed ledger", "crypto chain", "web3"],
        "tags": ["blockchain", "distributed ledger", "web3", "crypto", "chain", "decentralized", "technology", "transaction"],
        "primitives": [rect(2, 8, 6, 6), rect(9, 3, 6, 6), rect(16, 8, 6, 6), rect(9, 15, 6, 6), line(8, 10, 9, 8), line(15, 8, 16, 10), line(16, 13, 15, 16), line(9, 16, 8, 13)],
    },
    {
        "id": "iot",
        "name": "Internet of Things",
        "category": "Technology",
        "aliases": ["iot", "connected devices", "smart devices"],
        "tags": ["iot", "internet of things", "devices", "connected", "sensor", "smart", "network", "technology"],
        "primitives": [rect(8, 8, 8, 8), ellipse(10, 10, 4, 4, filled=True), line(12, 8, 12, 3), line(12, 16, 12, 21), line(8, 12, 3, 12), line(16, 12, 21, 12), ellipse(10.5, 1.5, 3, 3), ellipse(10.5, 19.5, 3, 3), ellipse(1.5, 10.5, 3, 3), ellipse(19.5, 10.5, 3, 3)],
    },
    {
        "id": "sensor",
        "name": "Sensor",
        "category": "Technology",
        "aliases": ["smart sensor", "telemetry device", "measurement sensor"],
        "tags": ["sensor", "telemetry", "measurement", "iot", "device", "signal", "monitoring", "technology"],
        "primitives": [ellipse(8, 8, 8, 8), ellipse(10.5, 10.5, 3, 3, filled=True), line(7, 7, 4, 4), line(17, 7, 20, 4), line(7, 17, 4, 20), line(17, 17, 20, 20), ellipse(5, 5, 14, 14)],
    },
    {
        "id": "digital-twin",
        "name": "Digital Twin",
        "category": "Technology",
        "aliases": ["virtual replica", "digital replica", "simulation twin"],
        "tags": ["digital twin", "virtual", "replica", "simulation", "model", "asset", "industry 4.0", "technology"],
        "primitives": [rect(2, 5, 8, 14), rect(14, 5, 8, 14), line(10, 8, 14, 8), line(10, 12, 14, 12), line(10, 16, 14, 16), ellipse(4.5, 8, 3, 3, filled=True), ellipse(16.5, 13, 3, 3, filled=True)],
    },
    {
        "id": "dashboard",
        "name": "Digital Dashboard",
        "category": "Technology",
        "aliases": ["control panel", "monitoring dashboard", "business dashboard"],
        "tags": ["dashboard", "control panel", "monitoring", "analytics", "kpi", "visualization", "software", "technology"],
        "primitives": [rect(2, 3, 20, 18), line(2, 8, 22, 8), rect(5, 11, 4, 7), rect(11, 14, 3, 4), rect(16, 10, 3, 8), ellipse(4, 5, 1.5, 1.5, filled=True)],
    },
    {
        "id": "git-branch",
        "name": "Code Branch",
        "category": "Technology",
        "aliases": ["git branch", "version control", "source branch"],
        "tags": ["git", "branch", "version control", "source", "code", "development", "merge", "repository"],
        "primitives": [ellipse(3, 2, 5, 5), ellipse(16, 2, 5, 5), ellipse(9.5, 17, 5, 5), line(5.5, 7, 5.5, 12), line(18.5, 7, 18.5, 10), line(5.5, 12, 12, 17), line(18.5, 10, 12, 17)],
    },
    {
        "id": "bug",
        "name": "Software Defect",
        "category": "Technology",
        "aliases": ["bug", "software issue", "defect"],
        "tags": ["bug", "defect", "software", "issue", "debug", "quality", "testing", "technology"],
        "primitives": [ellipse(7, 6, 10, 14), line(9, 6, 7, 3), line(15, 6, 17, 3), line(7, 10, 3, 8), line(7, 14, 3, 14), line(17, 10, 21, 8), line(17, 14, 21, 14), line(9, 10, 15, 10), line(12, 10, 12, 20)],
    },
    {
        "id": "browser",
        "name": "Web Application",
        "category": "Technology",
        "aliases": ["browser window", "web app", "website"],
        "tags": ["browser", "web application", "website", "internet", "software", "digital", "frontend", "technology"],
        "primitives": [rect(2, 3, 20, 18), line(2, 8, 22, 8), ellipse(4, 5, 1.5, 1.5, filled=True), ellipse(7, 5, 1.5, 1.5, filled=True), line(6, 13, 18, 13), line(6, 17, 14, 17)],
    },
    {
        "id": "data-lake",
        "name": "Data Lake",
        "category": "Technology",
        "aliases": ["enterprise data lake", "data repository", "lakehouse"],
        "tags": ["data lake", "lakehouse", "data", "storage", "analytics", "repository", "platform", "technology"],
        "primitives": [ellipse(4, 3, 16, 5), line(4, 5.5, 4, 13), line(20, 5.5, 20, 13), *polyline((3, 14), (7, 12), (11, 14), (15, 12), (21, 14)), *polyline((3, 18), (7, 16), (11, 18), (15, 16), (21, 18))],
    },
    {
        "id": "quantum",
        "name": "Quantum Computing",
        "category": "Technology",
        "aliases": ["quantum", "qubit", "quantum technology"],
        "tags": ["quantum", "computing", "qubit", "technology", "research", "innovation", "advanced", "processor"],
        "primitives": [ellipse(10, 10, 4, 4, filled=True), ellipse(3, 8, 18, 8), ellipse(8, 3, 8, 18), line(5, 5, 19, 19), line(19, 5, 5, 19)],
    },
    {
        "id": "machine-learning",
        "name": "Machine Learning Model",
        "category": "Technology",
        "aliases": ["ml model", "predictive model", "learning algorithm"],
        "tags": ["machine learning", "ml", "model", "algorithm", "prediction", "training", "artificial intelligence", "technology"],
        "primitives": [ellipse(2, 9, 5, 5, filled=True), ellipse(9.5, 2, 5, 5, filled=True), ellipse(17, 9, 5, 5, filled=True), ellipse(9.5, 17, 5, 5, filled=True), line(7, 11.5, 9.5, 4.5), line(14.5, 4.5, 17, 11.5), line(17, 11.5, 12, 19.5), line(12, 19.5, 7, 11.5), line(7, 11.5, 17, 11.5)],
    },
    {
        "id": "airplane",
        "name": "Air Transport",
        "category": "Operations",
        "aliases": ["airplane", "flight", "aviation"],
        "tags": ["airplane", "flight", "aviation", "transport", "travel", "logistics", "delivery", "operations"],
        "primitives": [line(12, 2, 12, 22), *polyline((12, 8), (3, 13), (3, 16), (12, 13)), *polyline((12, 8), (21, 13), (21, 16), (12, 13)), *polyline((12, 18), (7, 21), (12, 20), (17, 21), (12, 18))],
    },
    {
        "id": "ship",
        "name": "Maritime Transport",
        "category": "Operations",
        "aliases": ["cargo ship", "ocean freight", "vessel"],
        "tags": ["ship", "maritime", "ocean", "freight", "transport", "logistics", "cargo", "operations"],
        "primitives": [*polyline((2, 13), (22, 13), (19, 19), (12, 21), (5, 19), (2, 13), closed=True), rect(6, 8, 12, 5), rect(9, 4, 6, 4), line(2, 22, 7, 22), line(10, 22, 15, 22), line(18, 22, 22, 22)],
    },
    {
        "id": "conveyor",
        "name": "Production Line",
        "category": "Operations",
        "aliases": ["conveyor belt", "assembly line", "production conveyor"],
        "tags": ["conveyor", "production line", "assembly", "manufacturing", "process", "factory", "automation", "operations"],
        "primitives": [rect(3, 7, 5, 5), rect(10, 7, 5, 5), rect(17, 7, 5, 5), line(2, 14, 22, 14), ellipse(4, 16, 4, 4), ellipse(10, 16, 4, 4), ellipse(16, 16, 4, 4), line(6, 20, 18, 20)],
    },
    {
        "id": "inventory",
        "name": "Inventory",
        "category": "Operations",
        "aliases": ["stock", "inventory management", "stored goods"],
        "tags": ["inventory", "stock", "goods", "warehouse", "supply chain", "storage", "materials", "operations"],
        "primitives": [rect(3, 3, 8, 8), rect(13, 3, 8, 8), rect(3, 13, 8, 8), rect(13, 13, 8, 8), line(7, 3, 7, 11), line(17, 3, 17, 11), line(7, 13, 7, 21), line(17, 13, 17, 21)],
    },
    {
        "id": "procurement",
        "name": "Procurement",
        "category": "Operations",
        "aliases": ["purchasing", "source to pay", "buying"],
        "tags": ["procurement", "purchasing", "sourcing", "buying", "supplier", "purchase order", "cost", "operations"],
        "primitives": [rect(3, 4, 13, 16), line(6, 8, 13, 8), line(6, 12, 13, 12), line(6, 16, 11, 16), *polyline((15, 13), (18, 16), (22, 9)), line(16, 4, 21, 4), line(21, 4, 21, 8)],
    },
    {
        "id": "supplier",
        "name": "Supplier",
        "category": "Operations",
        "aliases": ["vendor", "supplier company", "external provider"],
        "tags": ["supplier", "vendor", "provider", "procurement", "partner", "supply chain", "external", "operations"],
        "primitives": [rect(3, 8, 8, 12), rect(13, 4, 8, 16), line(7, 4, 7, 8), line(4, 20, 22, 20), rect(5, 11, 2, 2, filled=True), rect(15, 8, 2, 2, filled=True), rect(18, 8, 2, 2, filled=True), line(11, 14, 13, 14)],
    },
    {
        "id": "quality",
        "name": "Quality Assurance",
        "category": "Operations",
        "aliases": ["quality control", "qa", "quality check"],
        "tags": ["quality", "assurance", "quality control", "qa", "inspection", "standard", "excellence", "operations"],
        "primitives": [ellipse(3, 3, 18, 18), *polyline((7, 12), (10.5, 15.5), (17.5, 8.5)), line(12, 3, 12, 6), line(12, 18, 12, 21), line(3, 12, 6, 12), line(18, 12, 21, 12)],
    },
    {
        "id": "map",
        "name": "Market Map",
        "category": "Operations",
        "aliases": ["geographic map", "territory map", "site map"],
        "tags": ["map", "geography", "market", "territory", "location", "region", "footprint", "operations"],
        "primitives": [*polyline((2, 5), (8, 2), (16, 5), (22, 2), (22, 19), (16, 22), (8, 19), (2, 22), (2, 5), closed=True), line(8, 2, 8, 19), line(16, 5, 16, 22)],
    },
    {
        "id": "barcode",
        "name": "Barcode",
        "category": "Operations",
        "aliases": ["product code", "sku code", "scan code"],
        "tags": ["barcode", "sku", "product", "scan", "inventory", "tracking", "retail", "operations"],
        "primitives": [rect(2, 4, 20, 16), line(5, 7, 5, 17), line(7, 7, 7, 17), line(10, 7, 10, 17), line(14, 7, 14, 17), line(16, 7, 16, 17), line(19, 7, 19, 17)],
    },
    {
        "id": "storefront",
        "name": "Retail Store",
        "category": "Operations",
        "aliases": ["storefront", "shop", "retail outlet"],
        "tags": ["store", "retail", "shop", "outlet", "channel", "sales", "location", "operations"],
        "primitives": [rect(3, 9, 18, 12), *polyline((3, 9), (5, 3), (19, 3), (21, 9)), line(3, 9, 21, 9), rect(6, 13, 6, 8), rect(15, 13, 3, 3), line(7, 6, 17, 6)],
    },
    {
        "id": "recruitment",
        "name": "Recruitment",
        "category": "People",
        "aliases": ["hiring", "talent acquisition", "candidate search"],
        "tags": ["recruitment", "hiring", "talent", "candidate", "workforce", "search", "human resources", "people"],
        "primitives": [ellipse(4, 3, 7, 7), *polyline((2, 19), (3, 14), (6, 11), (11, 11), (14, 14)), ellipse(13, 12, 7, 7), line(18, 18, 22, 22)],
    },
    {
        "id": "training",
        "name": "Training",
        "category": "People",
        "aliases": ["employee training", "learning program", "skills development"],
        "tags": ["training", "learning", "skills", "development", "education", "capability", "workforce", "people"],
        "primitives": [rect(3, 4, 18, 13), line(12, 17, 12, 21), line(8, 21, 16, 21), ellipse(6, 7, 4, 4), line(8, 11, 8, 15), *polyline((12, 8), (15, 11), (19, 6))],
    },
    {
        "id": "leadership",
        "name": "Leadership",
        "category": "People",
        "aliases": ["leader", "management team", "executive leadership"],
        "tags": ["leadership", "leader", "management", "executive", "direction", "team", "organization", "people"],
        "primitives": [ellipse(8, 2, 8, 8), *polyline((5, 20), (6, 14), (9, 11), (15, 11), (18, 14), (19, 20)), line(12, 11, 12, 18), *polyline((4, 8), (2, 5), (5, 3)), *polyline((20, 8), (22, 5), (19, 3))],
    },
    {
        "id": "diversity",
        "name": "Diversity and Inclusion",
        "category": "People",
        "aliases": ["diversity", "inclusion", "dei"],
        "tags": ["diversity", "inclusion", "dei", "belonging", "equity", "workforce", "culture", "people"],
        "primitives": [ellipse(2, 4, 5, 5), ellipse(9.5, 2, 5, 5), ellipse(17, 4, 5, 5), ellipse(5, 15, 5, 5), ellipse(14, 15, 5, 5), line(4.5, 9, 7.5, 15), line(19.5, 9, 16.5, 15), line(12, 7, 12, 15)],
    },
    {
        "id": "feedback",
        "name": "Feedback",
        "category": "People",
        "aliases": ["employee feedback", "review comment", "performance feedback"],
        "tags": ["feedback", "review", "comment", "performance", "conversation", "coaching", "employee", "people"],
        "primitives": [rect(3, 3, 18, 14), *polyline((8, 17), (7, 21), (12, 17)), line(7, 8, 17, 8), line(7, 12, 14, 12), ellipse(16, 11, 2, 2, filled=True)],
    },
    {
        "id": "remote-work",
        "name": "Remote Work",
        "category": "People",
        "aliases": ["work from home", "hybrid work", "distributed workforce"],
        "tags": ["remote work", "work from home", "hybrid", "distributed", "workforce", "home office", "digital workplace", "people"],
        "primitives": [rect(5, 8, 14, 10), *polyline((3, 9), (12, 2), (21, 9)), rect(9, 12, 6, 6), line(7, 22, 17, 22), line(12, 18, 12, 22), ellipse(10.5, 5, 3, 3, filled=True)],
    },
    {
        "id": "graduation",
        "name": "Capability Building",
        "category": "People",
        "aliases": ["graduation cap", "education", "learning academy"],
        "tags": ["capability building", "graduation", "education", "learning", "academy", "skills", "development", "people"],
        "primitives": [*polyline((2, 9), (12, 3), (22, 9), (12, 15), (2, 9), closed=True), *polyline((6, 12), (6, 17), (12, 21), (18, 17), (18, 12)), line(22, 9, 22, 17), ellipse(21, 17, 2, 2, filled=True)],
    },
    {
        "id": "identity",
        "name": "Identity Management",
        "category": "Security",
        "aliases": ["identity access management", "iam", "user identity"],
        "tags": ["identity", "iam", "access", "authentication", "user", "profile", "security", "governance"],
        "primitives": [rect(3, 4, 18, 16), ellipse(6, 7, 6, 6), *polyline((5, 18), (6, 15), (9, 13), (12, 15), (13, 18)), line(15, 8, 19, 8), line(15, 12, 19, 12), line(15, 16, 18, 16)],
    },
    {
        "id": "backup",
        "name": "Data Backup",
        "category": "Security",
        "aliases": ["backup copy", "data recovery", "restore point"],
        "tags": ["backup", "restore", "recovery", "data", "resilience", "continuity", "storage", "security"],
        "primitives": [ellipse(5, 6, 14, 14), *polyline((8, 5), (5, 6), (6, 9)), *polyline((16, 15), (19, 14), (18, 11)), line(5, 6, 11, 3), line(19, 14, 13, 19), rect(9, 8, 6, 8)],
    },
    {
        "id": "disaster-recovery",
        "name": "Disaster Recovery",
        "category": "Security",
        "aliases": ["business continuity", "dr plan", "service recovery"],
        "tags": ["disaster recovery", "business continuity", "recovery", "resilience", "outage", "restore", "risk", "security"],
        "primitives": [ellipse(3, 4, 18, 18), *polyline((4, 9), (7, 5), (11, 5)), line(7, 2, 7, 5), line(4, 5, 7, 5), *polyline((20, 15), (17, 20), (13, 20)), line(17, 20, 17, 23), line(20, 20, 17, 20), line(8, 15, 11, 10), line(11, 10, 14, 14), line(14, 14, 17, 9)],
    },
    {
        "id": "vulnerability",
        "name": "Vulnerability",
        "category": "Security",
        "aliases": ["security weakness", "exposure", "cyber vulnerability"],
        "tags": ["vulnerability", "weakness", "exposure", "cyber", "risk", "threat", "security", "remediation"],
        "primitives": [*polyline((12, 2), (20, 5), (19, 14), (16, 19), (12, 22), (8, 19), (5, 14), (4, 5), (12, 2), closed=True), line(12, 7, 10, 11), line(10, 11, 14, 14), line(14, 14, 11, 18)],
    },
    {
        "id": "risk-matrix",
        "name": "Risk Matrix",
        "category": "Security",
        "aliases": ["risk heatmap", "likelihood impact matrix", "risk assessment grid"],
        "tags": ["risk matrix", "heatmap", "likelihood", "impact", "assessment", "controls", "governance", "security"],
        "primitives": [rect(3, 3, 18, 18), line(9, 3, 9, 21), line(15, 3, 15, 21), line(3, 9, 21, 9), line(3, 15, 21, 15), ellipse(16.5, 4.5, 3, 3, filled=True), ellipse(10.5, 10.5, 3, 3, filled=True), ellipse(4.5, 16.5, 3, 3, filled=True)],
    },
    {
        "id": "headset",
        "name": "Customer Support",
        "category": "Communication",
        "aliases": ["headset", "service desk", "contact center"],
        "tags": ["headset", "support", "customer service", "contact center", "help desk", "call", "service", "communication"],
        "primitives": [ellipse(4, 3, 16, 16), rect(2, 10, 5, 8), rect(17, 10, 5, 8), line(17, 18, 15, 21), line(15, 21, 11, 21), ellipse(9.5, 19.5, 3, 3, filled=True)],
    },
    {
        "id": "paperclip",
        "name": "Attachment",
        "category": "Communication",
        "aliases": ["paperclip", "file attachment", "attached document"],
        "tags": ["attachment", "paperclip", "file", "document", "email", "message", "share", "communication"],
        "primitives": [*polyline((8, 7), (15, 3), (20, 5), (21, 10), (10, 21), (4, 19), (3, 14), (13, 5), (17, 7), (18, 10), (9, 18), (6, 17), (5, 14), (12, 8))],
    },
    {
        "id": "share",
        "name": "Share",
        "category": "Communication",
        "aliases": ["share content", "distribute", "social share"],
        "tags": ["share", "distribute", "send", "content", "collaboration", "network", "forward", "communication"],
        "primitives": [ellipse(2, 9, 6, 6), ellipse(16, 2, 6, 6), ellipse(16, 16, 6, 6), line(7.5, 10.5, 16.5, 6.5), line(7.5, 13.5, 16.5, 17.5)],
    },
    {
        "id": "tree",
        "name": "Nature",
        "category": "ESG",
        "aliases": ["tree", "forest", "biodiversity"],
        "tags": ["tree", "nature", "forest", "biodiversity", "environment", "climate", "ecology", "esg"],
        "primitives": [ellipse(5, 3, 14, 12), ellipse(2, 8, 10, 9), ellipse(12, 8, 10, 9), rect(10, 14, 4, 8), line(5, 22, 19, 22)],
    },
    {
        "id": "emissions",
        "name": "Carbon Emissions",
        "category": "ESG",
        "aliases": ["co2 emissions", "carbon footprint", "greenhouse gas"],
        "tags": ["emissions", "carbon", "co2", "greenhouse gas", "climate", "decarbonization", "footprint", "esg"],
        "primitives": [ellipse(2, 10, 8, 7), ellipse(7, 6, 9, 9), ellipse(14, 9, 8, 8), line(3, 17, 21, 17), rect(5, 17, 3, 5), rect(11, 17, 3, 5), rect(17, 17, 3, 5)],
    },
    {
        "id": "electric-vehicle",
        "name": "Electric Vehicle",
        "category": "ESG",
        "aliases": ["ev", "electric car", "clean transport"],
        "tags": ["electric vehicle", "ev", "car", "transport", "mobility", "electricity", "decarbonization", "esg"],
        "primitives": [*polyline((3, 15), (5, 9), (8, 6), (16, 6), (19, 9), (21, 15)), line(3, 15, 21, 15), ellipse(5, 14, 5, 5), ellipse(14, 14, 5, 5), *polyline((12, 7), (9, 12), (12, 12), (11, 16), (16, 10), (13, 10))],
    },
    {
        "id": "earth-leaf",
        "name": "Planet Positive",
        "category": "ESG",
        "aliases": ["green planet", "sustainable earth", "nature positive"],
        "tags": ["planet", "earth", "nature positive", "sustainability", "environment", "climate", "biodiversity", "esg"],
        "primitives": [ellipse(2, 2, 18, 18), line(2, 11, 20, 11), ellipse(7, 2, 8, 18), *polyline((13, 20), (16, 15), (22, 13), (20, 19), (16, 22), (13, 20), closed=True), line(15, 21, 20, 16)],
    },
    {
        "id": "waste",
        "name": "Waste Management",
        "category": "ESG",
        "aliases": ["waste bin", "resource recovery", "material waste"],
        "tags": ["waste", "waste management", "materials", "recycling", "resource recovery", "environment", "circularity", "esg"],
        "primitives": [*polyline((6, 7), (18, 7), (17, 21), (7, 21), (6, 7), closed=True), line(4, 7, 20, 7), line(9, 4, 15, 4), line(10, 10, 10, 18), line(14, 10, 14, 18)],
    },
]

ICONS.extend(CONSULTING_ICONS)

FAMILY_IDS = [
    "document", "folder", "user", "users", "cloud", "database", "server", "shield", "lock",
    "analytics", "calendar", "chat", "mail", "globe", "target", "lightbulb", "package", "truck",
    "factory", "building", "currency", "wallet", "bank", "chip", "network", "automation",
    "checklist", "presentation", "phone", "code", "sustainability", "energy", "risk", "process",
]

for family_id in FAMILY_IDS:
    base = next(icon for icon in ICONS if icon["id"] == family_id)
    for badge_id, variant_name, variant_terms in VARIANTS:
        icon_id = f"{family_id}-{badge_id}"
        alias_roots = base.get("aliases", ALIASES_BY_ID.get(family_id, [base["name"].lower()]))
        ICONS.append(
            {
                "id": icon_id,
                "name": f"{base['name']} {variant_name}",
                "category": base["category"],
                "aliases": [
                    f"{alias_roots[0]} {variant_terms[0]}",
                    f"{base['name'].lower()} {variant_terms[1]}",
                    f"{variant_terms[2]} {alias_roots[-1]}",
                ],
                "tags": list(dict.fromkeys([*base["tags"], *variant_terms, badge_id, variant_name.lower()])),
                "primitives": [*transformed(base["primitives"], 0.72, 0.5, 0.5), *BADGES[badge_id]],
            }
        )

for icon in ICONS:
    if "aliases" not in icon:
        icon["aliases"] = ALIASES_BY_ID[icon["id"]]
    icon["tags"].extend(EXTRA_TAGS_BY_ID.get(icon["id"], []))

BENCHMARKS = [
    {
        "name": "Tabler Icons",
        "license": "MIT",
        "url": "https://tabler.io/icons",
        "takeaways": ["24x24 grid", "2px outline style", "large category-filtered catalog", "color and stroke customization"],
    },
    {
        "name": "Lucide",
        "license": "ISC with some MIT-derived icons",
        "url": "https://lucide.dev/",
        "takeaways": ["consistent community-maintained outline language", "simple semantic names", "24x24 compatibility"],
    },
    {
        "name": "Phosphor Icons",
        "license": "MIT",
        "url": "https://phosphoricons.com/",
        "takeaways": ["first-class tags and categories", "aliases", "presentation-friendly breadth", "multiple weights"],
    },
    {
        "name": "Material Symbols",
        "license": "Apache-2.0",
        "url": "https://fonts.google.com/icons",
        "takeaways": ["broad concept coverage", "optical sizing", "variable fill and weight"],
    },
    {
        "name": "Heroicons",
        "license": "MIT",
        "url": "https://heroicons.com/",
        "takeaways": ["careful optical balance", "clean 24x24 outline set", "strong small-set consistency"],
    },
]


def validate_catalog(icons: list[dict[str, Any]]) -> None:
    seen: set[str] = set()
    allowed = {"line", "rect", "ellipse"}
    coordinate_keys = {"x", "y", "x1", "y1", "x2", "y2", "width", "height"}
    for icon in icons:
        icon_id = icon.get("id")
        if not isinstance(icon_id, str) or not icon_id.replace("-", "").isalnum():
            raise ValueError(f"Invalid icon id: {icon_id!r}")
        if icon_id in seen:
            raise ValueError(f"Duplicate icon id: {icon_id}")
        seen.add(icon_id)
        if not icon.get("name") or not icon.get("category") or not icon.get("tags") or not icon.get("aliases"):
            raise ValueError(f"Icon {icon_id} is missing searchable metadata")
        if len(icon["tags"]) < 8 or len(icon["aliases"]) < 2:
            raise ValueError(f"Icon {icon_id} needs more search terms")
        if len(set(icon["tags"])) != len(icon["tags"]) or len(set(icon["aliases"])) != len(icon["aliases"]):
            raise ValueError(f"Icon {icon_id} has duplicate search terms")
        primitives = icon.get("primitives")
        if not isinstance(primitives, list) or len(primitives) < 2:
            raise ValueError(f"Icon {icon_id} must contain at least two primitives")
        for primitive in primitives:
            if primitive.get("kind") not in allowed:
                raise ValueError(f"Icon {icon_id} has unsupported primitive {primitive.get('kind')!r}")
            for key in coordinate_keys & primitive.keys():
                value = primitive[key]
                if not isinstance(value, (int, float)) or value < 0 or value > 24:
                    raise ValueError(f"Icon {icon_id} has invalid {key}: {value!r}")
            if primitive.get("width", 1) <= 0 or primitive.get("height", 1) <= 0:
                raise ValueError(f"Icon {icon_id} has a non-positive primitive size")


def catalog_json() -> str:
    validate_catalog(ICONS)
    catalog = {
        "schema": 2,
        "viewBox": 24,
        "style": {"name": "IconAid Outline", "stroke": 1.5, "lineCap": "round", "lineJoin": "round"},
        "provenance": "Original SlideAid geometry informed by the catalog benchmark; no third-party paths are copied.",
        "license": "MIT; see shared/iconaid/LICENSE",
        "benchmarks": BENCHMARKS,
        "icons": ICONS,
    }
    return json.dumps(catalog, indent=2, sort_keys=True) + "\n"


def vba_string(value: str) -> str:
    return '"' + value.replace('"', '""') + '"'


def vba_number(value: float) -> str:
    return f"{value:g}!"


def primitive_data(icon: dict[str, Any]) -> str:
    commands = []
    for primitive in icon["primitives"]:
        if primitive["kind"] == "line":
            values = [primitive[key] for key in ("x1", "y1", "x2", "y2")]
            commands.append("L," + ",".join(f"{value:g}" for value in values))
        else:
            values = [primitive[key] for key in ("x", "y", "width", "height")]
            prefix = "R" if primitive["kind"] == "rect" else "E"
            commands.append(prefix + "," + ",".join(f"{value:g}" for value in values) + f",{int(primitive.get('filled', False))}")
    return ";".join(commands)


def icon_record(icon: dict[str, Any]) -> str:
    search_text = " ".join([icon["name"], icon["category"], *icon["aliases"], *icon["tags"]])
    return "|".join([icon["name"], icon["category"], search_text, primitive_data(icon)])


def render_vba_controller(chunk_count: int) -> str:
    lines = [
        'Attribute VB_Name = "modIconAid"',
        "' Generated by scripts/build_icon_catalog.py. Do not edit by hand.",
        "Option Explicit",
        "",
        "Private Const IA_SIZE As Single = 72!",
        "Private Const IA_VIEWBOX As Single = 24!",
        "Private Const IA_MAX_RESULTS As Long = 12",
        "",
        "Public Sub SearchAndInsertIcon()",
        "    Dim query As String, answer As String",
        "    Dim matches() As Long, matchCount As Long, totalCount As Long",
        "    Dim i As Long, selected As Long",
        "    Dim sourceSlide As Slide, previewSlide As Slide",
        "    Dim insertedIcon As Shape",
        "    Dim errorNumber As Long, errorText As String",
        "    On Error GoTo Failed",
        "    Set sourceSlide = CurrentSlide()",
        '    query = Trim$(InputBox("Search by name, category, alias, or keyword. Leave blank to browse the first results.", "IconAid"))',
        "    For i = 1 To IA_IconCount()",
        "        If Len(query) = 0 Or InStr(1, IA_RecordField(i, 2), query, vbTextCompare) > 0 Then",
        "            totalCount = totalCount + 1",
        "            If matchCount < IA_MAX_RESULTS Then",
        "                matchCount = matchCount + 1",
        "                ReDim Preserve matches(1 To matchCount)",
        "                matches(matchCount) = i",
        "            End If",
        "        End If",
        "    Next i",
        "    If totalCount = 0 Then",
        '        MsgBox "No icons match """ & query & """.", vbInformation, "IconAid"',
        "        Exit Sub",
        "    End If",
        "    Set previewSlide = ActivePresentation.Slides.Add(ActivePresentation.Slides.Count + 1, ppLayoutBlank)",
        "    Call IA_DrawPreview(previewSlide, matches, matchCount, totalCount, query)",
        "    ActiveWindow.View.GotoSlide previewSlide.SlideIndex",
        "    DoEvents",
        '    answer = Trim$(InputBox("Enter an icon number shown on the slide. Cancel returns without inserting.", "IconAid"))',
        "    If Len(answer) = 0 Then GoTo Cancelled",
        "    If Not IsNumeric(answer) Then",
        '        MsgBox "Enter a number shown in the visual results.", vbInformation, "IconAid"',
        "        GoTo Cancelled",
        "    End If",
        "    selected = CLng(answer)",
        "    If selected < 1 Or selected > matchCount Then",
        '        MsgBox "The icon number is outside the visual results.", vbInformation, "IconAid"',
        "        GoTo Cancelled",
        "    End If",
        "    previewSlide.Delete",
        "    Set previewSlide = Nothing",
        "    ActiveWindow.View.GotoSlide sourceSlide.SlideIndex",
        "    Set insertedIcon = IA_RenderIcon(sourceSlide, matches(selected), _",
        "        (ActivePresentation.PageSetup.SlideWidth - IA_SIZE) / 2, _",
        "        (ActivePresentation.PageSetup.SlideHeight - IA_SIZE) / 2, _",
        "        IA_SIZE, RGB(31, 73, 125))",
        "    insertedIcon.Select",
        "    Exit Sub",
        "Cancelled:",
        "    On Error Resume Next",
        "    If Not previewSlide Is Nothing Then previewSlide.Delete",
        "    ActiveWindow.View.GotoSlide sourceSlide.SlideIndex",
        "    On Error GoTo 0",
        "    Exit Sub",
        "Failed:",
        "    errorNumber = Err.Number",
        "    errorText = Err.Description",
        "    On Error Resume Next",
        "    If Not previewSlide Is Nothing Then previewSlide.Delete",
        "    If Not sourceSlide Is Nothing Then ActiveWindow.View.GotoSlide sourceSlide.SlideIndex",
        "    On Error GoTo 0",
        '    MsgBox "IconAid could not complete the operation: " & errorText, vbExclamation, "IconAid"',
        "End Sub",
        "",
        "Private Function IA_IconCount() As Long",
        f"    IA_IconCount = {len(ICONS)}",
        "End Function",
        "",
        "Private Function IA_Record(ByVal index As Long) As String",
        "    Select Case index",
    ]
    for chunk_index in range(1, chunk_count + 1):
        start = (chunk_index - 1) * 60 + 1
        end = min(chunk_index * 60, len(ICONS))
        lines.append(f"        Case {start} To {end}: IA_Record = IA_Data{chunk_index:02d}(index)")
    lines.extend(
        [
            "        Case Else",
            '            Err.Raise vbObjectError + 742, "IconAid", "Unknown icon."',
            "    End Select",
            "End Function",
            "",
            "Private Function IA_RecordField(ByVal index As Long, ByVal fieldIndex As Long) As String",
            "    Dim fields() As String",
            "    fields = Split(IA_Record(index), \"|\")",
            "    IA_RecordField = fields(fieldIndex)",
            "End Function",
            "",
            "Private Sub IA_DrawPreview(ByVal previewSlide As Slide, ByRef matches() As Long, _",
            "        ByVal matchCount As Long, ByVal totalCount As Long, ByVal query As String)",
            "    Dim slideWidth As Single, slideHeight As Single",
            "    Dim cellWidth As Single, cellHeight As Single, iconSize As Single",
            "    Dim rowIndex As Long, columnIndex As Long, i As Long",
            "    Dim originLeft As Single, originTop As Single",
            "    Dim titleShape As Shape, labelShape As Shape, previewIcon As Shape",
            "    slideWidth = ActivePresentation.PageSetup.SlideWidth",
            "    slideHeight = ActivePresentation.PageSetup.SlideHeight",
            "    cellWidth = (slideWidth - 48!) / 4!",
            "    cellHeight = (slideHeight - 92!) / 3!",
            "    iconSize = IA_Min(54!, cellHeight - 34!)",
            "    Set titleShape = previewSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 24!, 14!, slideWidth - 48!, 28!)",
            '    titleShape.TextFrame.TextRange.Text = "IconAid results" & IIf(Len(query) > 0, " for """ & query & """", "")',
            "    titleShape.TextFrame.TextRange.Font.Size = 18!",
            "    titleShape.TextFrame.TextRange.Font.Bold = msoTrue",
            "    titleShape.Line.Visible = msoFalse",
            "    titleShape.Fill.Visible = msoFalse",
            "    For i = 1 To matchCount",
            "        rowIndex = (i - 1) \\ 4",
            "        columnIndex = (i - 1) Mod 4",
            "        originLeft = 24! + columnIndex * cellWidth + (cellWidth - iconSize) / 2!",
            "        originTop = 48! + rowIndex * cellHeight + 4!",
            "        Set previewIcon = IA_RenderIcon(previewSlide, matches(i), originLeft, originTop, iconSize, RGB(31, 73, 125))",
            "        Set labelShape = previewSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, _",
            "            28! + columnIndex * cellWidth, originTop + iconSize + 4!, cellWidth - 8!, 24!)",
            '        labelShape.TextFrame.TextRange.Text = CStr(i) & ". " & IA_RecordField(matches(i), 0)',
            "        labelShape.TextFrame.TextRange.Font.Size = 10!",
            "        labelShape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter",
            "        labelShape.TextFrame.MarginLeft = 0!",
            "        labelShape.TextFrame.MarginRight = 0!",
            "        labelShape.Line.Visible = msoFalse",
            "        labelShape.Fill.Visible = msoFalse",
            "    Next i",
            "    If totalCount > matchCount Then",
            "        Set labelShape = previewSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 24!, slideHeight - 25!, slideWidth - 48!, 16!)",
            '        labelShape.TextFrame.TextRange.Text = CStr(totalCount - matchCount) & " more matches. Refine the search to see them."',
            "        labelShape.TextFrame.TextRange.Font.Size = 9!",
            "        labelShape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter",
            "        labelShape.Line.Visible = msoFalse",
            "        labelShape.Fill.Visible = msoFalse",
            "    End If",
            "End Sub",
            "",
            "Private Function IA_Min(ByVal firstValue As Single, ByVal secondValue As Single) As Single",
            "    If firstValue < secondValue Then IA_Min = firstValue Else IA_Min = secondValue",
            "End Function",
            "",
            "Private Function IA_RenderIcon(ByVal targetSlide As Slide, ByVal iconIndex As Long, _",
            "        ByVal originLeft As Single, ByVal originTop As Single, ByVal iconSize As Single, _",
            "        ByVal iconColor As Long) As Shape",
            "    Dim itemNames() As String, itemCount As Long, groupShape As Shape",
            "    Dim recordFields() As String, commands() As String, fields() As String",
            "    Dim i As Long, shapeType As Long",
            "    Dim errorNumber As Long, errorText As String",
            "    On Error GoTo Failed",
            "    recordFields = Split(IA_Record(iconIndex), \"|\")",
            "    commands = Split(recordFields(3), \";\")",
            "    For i = LBound(commands) To UBound(commands)",
            "        fields = Split(commands(i), \",\")",
            "        If fields(0) = \"L\" Then",
            "            Call IA_AddLine(targetSlide, itemNames, itemCount, originLeft, originTop, iconSize, iconColor, _",
            "                CSng(Val(fields(1))), CSng(Val(fields(2))), CSng(Val(fields(3))), CSng(Val(fields(4))))",
            "        Else",
            "            If fields(0) = \"R\" Then shapeType = msoShapeRectangle Else shapeType = msoShapeOval",
            "            Call IA_AddShape(targetSlide, itemNames, itemCount, originLeft, originTop, iconSize, iconColor, shapeType, _",
            "                CSng(Val(fields(1))), CSng(Val(fields(2))), CSng(Val(fields(3))), CSng(Val(fields(4))), CStr(fields(5)) = \"1\")",
            "        End If",
            "    Next i",
            "    Set groupShape = targetSlide.Shapes.Range(itemNames).Group",
            '    groupShape.AlternativeText = "IconAid vector icon: " & recordFields(0)',
            "    Set IA_RenderIcon = groupShape",
            "    Exit Function",
            "Failed:",
            "    errorNumber = Err.Number",
            "    errorText = Err.Description",
            "    On Error Resume Next",
            "    For i = itemCount To 1 Step -1",
            "        targetSlide.Shapes(itemNames(i)).Delete",
            "    Next i",
            "    On Error GoTo 0",
            '    Err.Raise errorNumber, "IconAid", errorText',
            "End Function",
            "",
            "Private Sub IA_AddLine(ByVal sl As Slide, ByRef names() As String, ByRef count As Long, _",
            "        ByVal baseLeft As Single, ByVal baseTop As Single, ByVal iconSize As Single, ByVal iconColor As Long, _",
            "        ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)",
            "    Dim item As Shape, scale As Single",
            "    scale = iconSize / IA_VIEWBOX",
            "    Set item = sl.Shapes.AddLine(baseLeft + x1 * scale, baseTop + y1 * scale, _",
            "                                 baseLeft + x2 * scale, baseTop + y2 * scale)",
            "    item.Line.ForeColor.RGB = iconColor",
            "    item.Line.Weight = 1.5",
            "    Call IA_RecordName(names, count, item.Name)",
            "End Sub",
            "",
            "Private Sub IA_AddShape(ByVal sl As Slide, ByRef names() As String, ByRef count As Long, _",
            "        ByVal baseLeft As Single, ByVal baseTop As Single, ByVal iconSize As Single, ByVal iconColor As Long, _",
            "        ByVal shapeType As Long, ByVal x As Single, ByVal y As Single, _",
            "        ByVal itemWidth As Single, ByVal itemHeight As Single, ByVal isFilled As Boolean)",
            "    Dim item As Shape, scale As Single",
            "    scale = iconSize / IA_VIEWBOX",
            "    Set item = sl.Shapes.AddShape(shapeType, baseLeft + x * scale, baseTop + y * scale, _",
            "                                  itemWidth * scale, itemHeight * scale)",
            "    If isFilled Then",
            "        item.Fill.Solid",
            "        item.Fill.ForeColor.RGB = iconColor",
            "        item.Line.Visible = msoFalse",
            "    Else",
            "        item.Fill.Visible = msoFalse",
            "        item.Line.ForeColor.RGB = iconColor",
            "        item.Line.Weight = 1.5",
            "    End If",
            "    Call IA_RecordName(names, count, item.Name)",
            "End Sub",
            "",
            "Private Sub IA_RecordName(ByRef names() As String, ByRef count As Long, ByVal itemName As String)",
            "    count = count + 1",
            "    ReDim Preserve names(1 To count)",
            "    names(count) = itemName",
            "End Sub",
            "",
        ]
    )
    return "\n".join(lines)


def render_vba_chunk(chunk_index: int, start_index: int, icons: list[dict[str, Any]]) -> str:
    function_name = f"IA_Data{chunk_index:02d}"
    lines = [
        f'Attribute VB_Name = "modIconAidData{chunk_index:02d}"',
        "' Generated by scripts/build_icon_catalog.py. Do not edit by hand.",
        "Option Explicit",
        "",
        f"Public Function {function_name}(ByVal index As Long) As String",
        "    Select Case index",
    ]
    for offset, icon in enumerate(icons):
        lines.append(f"        Case {start_index + offset}: {function_name} = {vba_string(icon_record(icon))}")
    lines.extend(["    End Select", "End Function", ""])
    return "\n".join(lines)


def render_vba_taskpane_bridge() -> str:
    return "\n".join(
        [
            'Attribute VB_Name = "modIconAid"',
            "' Generated by scripts/build_icon_catalog.py. Do not edit by hand.",
            "Option Explicit",
            "",
            "' IconAid uses an Office.js task pane. Keep this compatibility entry point",
            "' so older RibbonX packages compile without loading the retired VBA browser.",
            "Public Sub SearchAndInsertIcon()",
            '    MsgBox "Open IconAid from Home > Add-ins.", vbInformation, "IconAid"',
            "End Sub",
            "",
        ]
    )


def outputs() -> dict[Path, str]:
    validate_catalog(ICONS)
    chunk_size = 60
    chunks = [ICONS[start : start + chunk_size] for start in range(0, len(ICONS), chunk_size)]
    generated = {
        CATALOG_PATH: catalog_json(),
        VBA_PATH: render_vba_taskpane_bridge(),
    }
    for chunk_index, chunk in enumerate(chunks, 1):
        path = VBA_PATH.with_name(f"modIconAidData{chunk_index:02d}.bas")
        generated[path] = render_vba_chunk(chunk_index, (chunk_index - 1) * chunk_size + 1, chunk)
    return generated


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--check", action="store_true", help="Fail if generated files are stale.")
    args = parser.parse_args()
    stale: list[Path] = []
    for path, content in outputs().items():
        if args.check:
            if not path.exists() or path.read_text() != content:
                stale.append(path)
        else:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(content)
            print(path.relative_to(ROOT))
    if stale:
        names = ", ".join(str(path.relative_to(ROOT)) for path in stale)
        raise SystemExit(f"Generated IconAid files are stale: {names}")


if __name__ == "__main__":
    main()
