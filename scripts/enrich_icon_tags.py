#!/usr/bin/env python3
"""
Enrich icon tags with semantic keywords for better searchability.

This script adds rich, contextual tags to icons based on:
1. Icon name analysis (word stems, synonyms)
2. Category semantics
3. Business/consulting context
4. Visual concept mapping

Usage:
    python3 scripts/enrich_icon_tags.py
    python3 scripts/enrich_icon_tags.py --preview  # Show changes without saving
    python3 scripts/enrich_icon_tags.py --source tabler  # Only enrich one source

Output:
    Updates the combined-search-index.json with enriched tags
"""
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
INDEX_FILE = ROOT / "shared" / "iconaid" / "external-sources" / "combined-search-index.json"

# Semantic keyword mappings: pattern -> additional tags
# These add business/consulting context to common icon concepts

SEMANTIC_ENRICHMENTS = {
    # Analytics & Data
    r"chart|graph|analytics|statistic": [
        "metrics", "reporting", "dashboard", "kpi", "performance",
        "visualization", "data", "insights", "analysis", "measure"
    ],
    r"bar[-\s]?chart|column[-\s]?chart": [
        "comparison", "ranking", "distribution", "categories"
    ],
    r"line[-\s]?chart|trend": [
        "time series", "progression", "growth", "forecast", "historical"
    ],
    r"pie[-\s]?chart|donut": [
        "proportion", "share", "percentage", "breakdown", "composition"
    ],
    r"funnel": [
        "conversion", "pipeline", "sales funnel", "leads", "stages"
    ],
    
    # Business Concepts
    r"briefcase|portfolio": [
        "work", "business", "professional", "career", "job", "corporate",
        "consulting", "enterprise"
    ],
    r"target|bullseye|goal": [
        "objective", "aim", "strategy", "kpi", "north star", "focus",
        "achievement", "success"
    ],
    r"flag|milestone": [
        "achievement", "checkpoint", "progress", "completion", "deadline",
        "launch", "release", "delivery"
    ],
    r"trophy|award|medal": [
        "achievement", "success", "winner", "recognition", "excellence",
        "best practice"
    ],
    r"lightbulb|idea": [
        "innovation", "insight", "creative", "brainstorm", "solution",
        "concept", "inspiration", "thinking"
    ],
    r"puzzle": [
        "solution", "fit", "integration", "problem solving", "complexity",
        "strategy", "synergy", "components"
    ],
    r"handshake|partnership": [
        "deal", "agreement", "collaboration", "alliance", "contract",
        "cooperation", "joint venture", "merger"
    ],
    r"roadmap|route": [
        "strategy", "planning", "journey", "path", "direction", "timeline",
        "phases", "milestones"
    ],
    
    # Finance
    r"dollar|currency|money|cash": [
        "finance", "payment", "cost", "revenue", "budget", "investment",
        "funding", "capital", "expense", "price"
    ],
    r"bank|treasury": [
        "financial institution", "capital", "funding", "loan", "credit"
    ],
    r"wallet|payment": [
        "transaction", "purchase", "billing", "invoice", "checkout"
    ],
    r"calculator|calculate": [
        "math", "compute", "estimate", "forecast", "budget", "accounting"
    ],
    r"receipt|invoice": [
        "billing", "payment", "transaction", "purchase", "accounts"
    ],
    r"percent|percentage": [
        "ratio", "rate", "margin", "share", "growth rate", "discount"
    ],
    
    # Technology
    r"cloud": [
        "saas", "hosting", "infrastructure", "platform", "storage", 
        "internet", "online", "aws", "azure", "digital"
    ],
    r"server|data[-\s]?center": [
        "hosting", "infrastructure", "backend", "compute", "hardware"
    ],
    r"database|storage": [
        "data", "sql", "records", "warehouse", "repository", "store"
    ],
    r"api|integration|webhook": [
        "interface", "connection", "service", "endpoint", "microservice"
    ],
    r"code|programming|terminal": [
        "development", "software", "developer", "engineering", "tech"
    ],
    r"chip|processor|cpu": [
        "ai", "computing", "hardware", "semiconductor", "machine learning"
    ],
    r"network|nodes": [
        "connections", "topology", "distributed", "mesh", "ecosystem"
    ],
    r"robot|automation": [
        "rpa", "workflow", "automatic", "efficiency", "orchestration"
    ],
    
    # Security
    r"lock|secure": [
        "security", "protection", "privacy", "access control", "encryption"
    ],
    r"shield": [
        "security", "protection", "defense", "safe", "resilience", "guard"
    ],
    r"key|password": [
        "access", "authentication", "unlock", "credential", "login"
    ],
    r"eye|visibility": [
        "view", "monitor", "observe", "watch", "visibility", "insight"
    ],
    r"fingerprint|biometric": [
        "identity", "authentication", "verification", "secure"
    ],
    
    # Communication
    r"mail|email|envelope": [
        "message", "inbox", "correspondence", "communication", "send",
        "newsletter", "notification"
    ],
    r"chat|message|conversation": [
        "discussion", "talk", "communication", "dialogue", "feedback",
        "support", "instant message"
    ],
    r"phone|call|telephone": [
        "contact", "voice", "communication", "customer service", "call center"
    ],
    r"video|camera": [
        "meeting", "conference", "webinar", "recording", "streaming"
    ],
    r"microphone|audio": [
        "voice", "recording", "podcast", "speech", "sound"
    ],
    r"bell|notification|alert": [
        "reminder", "announcement", "warning", "attention", "update"
    ],
    r"megaphone|broadcast": [
        "announcement", "marketing", "campaign", "promotion", "news"
    ],
    
    # Documents
    r"file|document": [
        "paper", "attachment", "record", "report", "content"
    ],
    r"folder|directory": [
        "organize", "storage", "files", "collection", "archive"
    ],
    r"clipboard": [
        "copy", "paste", "notes", "list", "tasks"
    ],
    r"book|manual|guide": [
        "documentation", "knowledge", "reference", "learning", "reading"
    ],
    r"presentation|slides": [
        "powerpoint", "deck", "pitch", "slideshow", "keynote", "meeting"
    ],
    
    # People & Organization
    r"user|person|people": [
        "employee", "customer", "stakeholder", "individual", "human",
        "staff", "team member", "workforce"
    ],
    r"team|group|users": [
        "organization", "workforce", "collaboration", "department", "squad"
    ],
    r"building|office": [
        "company", "corporate", "headquarters", "workplace", "enterprise"
    ],
    r"hierarchy|org[-\s]?chart": [
        "structure", "organization", "reporting", "management", "leadership"
    ],
    
    # Operations & Logistics
    r"truck|delivery|shipping": [
        "logistics", "transport", "freight", "supply chain", "distribution"
    ],
    r"warehouse|inventory": [
        "storage", "stock", "fulfillment", "distribution center"
    ],
    r"factory|manufacturing": [
        "production", "industry", "plant", "operations", "output"
    ],
    r"box|package|parcel": [
        "product", "shipping", "delivery", "order", "goods"
    ],
    r"process|workflow|flow": [
        "procedure", "steps", "pipeline", "sequence", "value chain"
    ],
    r"checklist|tasks": [
        "todo", "quality", "audit", "completion", "requirements"
    ],
    r"wrench|tool|maintenance": [
        "repair", "service", "fix", "support", "engineering"
    ],
    
    # Time & Planning
    r"calendar|schedule": [
        "planning", "date", "event", "appointment", "timeline", "agenda"
    ],
    r"clock|time": [
        "duration", "deadline", "schedule", "wait", "history", "speed"
    ],
    r"hourglass|timer": [
        "deadline", "countdown", "waiting", "progress", "time limit"
    ],
    
    # Nature & ESG
    r"leaf|plant|tree": [
        "nature", "green", "environment", "sustainability", "ecology",
        "organic", "growth", "natural"
    ],
    r"sun|solar": [
        "energy", "renewable", "power", "bright", "day", "sustainable"
    ],
    r"wind|turbine": [
        "energy", "renewable", "power", "sustainable", "clean"
    ],
    r"water|droplet|wave": [
        "liquid", "hydro", "resource", "clean", "flow"
    ],
    r"recycle|circular": [
        "sustainability", "reuse", "environment", "eco", "waste reduction"
    ],
    r"battery|power": [
        "energy", "charge", "electric", "storage", "capacity"
    ],
    r"globe|earth|world": [
        "global", "international", "worldwide", "planet", "geography"
    ],
    
    # Arrows & Navigation
    r"arrow[-\s]?up|trend[-\s]?up|increase": [
        "growth", "improvement", "rise", "upward", "positive", "gain"
    ],
    r"arrow[-\s]?down|trend[-\s]?down|decrease": [
        "decline", "reduction", "fall", "downward", "negative", "loss"
    ],
    r"refresh|sync|reload": [
        "update", "synchronize", "renew", "repeat", "cycle"
    ],
    r"expand|maximize": [
        "enlarge", "grow", "scale up", "full screen"
    ],
    r"compress|minimize": [
        "shrink", "reduce", "scale down", "compact"
    ],
    
    # Status & States
    r"check|success|approved": [
        "complete", "verified", "done", "confirmed", "accepted", "valid"
    ],
    r"x[-\s]?mark|cancel|close|remove": [
        "delete", "reject", "failed", "blocked", "stop", "invalid"
    ],
    r"warning|exclamation|alert": [
        "caution", "attention", "issue", "problem", "notice"
    ],
    r"info|information": [
        "details", "help", "about", "documentation", "guidance"
    ],
    r"question|help": [
        "support", "faq", "assistance", "inquiry", "unknown"
    ],
    r"plus|add|new": [
        "create", "insert", "more", "increase", "expand"
    ],
    r"minus|subtract|remove": [
        "delete", "reduce", "less", "decrease"
    ],
    
    # Shapes & UI
    r"grid|layout": [
        "arrangement", "structure", "organization", "display", "view"
    ],
    r"list|menu": [
        "items", "options", "navigation", "selection", "choices"
    ],
    r"layers|stack": [
        "levels", "hierarchy", "architecture", "components", "tiers"
    ],
    r"filter|funnel": [
        "sort", "narrow", "refine", "search", "criteria"
    ],
    r"search|magnifying": [
        "find", "lookup", "discover", "explore", "query"
    ],
    r"settings|gear|cog": [
        "configuration", "preferences", "options", "customize", "admin"
    ],
    
    # Specific Animals (for visual icons)
    r"\bdog\b": [
        "pet", "animal", "canine", "puppy", "companion", "loyal"
    ],
    r"\bcat\b": [
        "pet", "animal", "feline", "kitten"
    ],
    r"\bbird\b": [
        "animal", "flying", "nature", "freedom", "twitter"
    ],
    r"\bfish\b": [
        "animal", "aquatic", "ocean", "sea", "swimming"
    ],
    r"\bbug\b|insect": [
        "software bug", "debug", "error", "issue", "problem"
    ],
    
    # Extended Animals & Creatures
    r"\bhorse\b|equestrian": [
        "animal", "riding", "race", "stable", "stallion", "pony", "equine"
    ],
    r"\bcow\b|cattle|bull": [
        "animal", "farm", "livestock", "agriculture", "dairy", "beef"
    ],
    r"\bpig\b|piggy": [
        "animal", "farm", "savings", "piggy bank", "oink"
    ],
    r"\brabbit\b|bunny": [
        "animal", "pet", "fast", "quick", "hop", "easter"
    ],
    r"\bbear\b": [
        "animal", "stock market", "bearish", "decline", "wild"
    ],
    r"\bbull\b": [
        "animal", "stock market", "bullish", "growth", "strong"
    ],
    r"\bwolf\b": [
        "animal", "wild", "pack", "team", "predator"
    ],
    r"\blion\b": [
        "animal", "king", "leader", "brave", "courage", "strength"
    ],
    r"\beagle\b": [
        "animal", "bird", "freedom", "vision", "america", "soar"
    ],
    r"\bdolphin\b": [
        "animal", "marine", "intelligent", "playful", "ocean"
    ],
    r"\bbutterfly\b": [
        "animal", "insect", "transformation", "change", "metamorphosis", "beauty"
    ],
    r"\bbee\b": [
        "animal", "insect", "busy", "productive", "honey", "hive", "teamwork"
    ],
    r"\bant\b": [
        "animal", "insect", "hardworking", "colony", "teamwork", "small"
    ],
    r"\bspider\b": [
        "animal", "insect", "web", "network", "crawl"
    ],
    r"\bsnake\b|serpent": [
        "animal", "reptile", "danger", "caution", "medical", "pharmacy"
    ],
    r"\bturtle\b|tortoise": [
        "animal", "slow", "steady", "patience", "shell", "protection"
    ],
    r"\bwhale\b": [
        "animal", "marine", "big", "large", "ocean", "crypto whale"
    ],
    r"\bshark\b": [
        "animal", "marine", "predator", "aggressive", "business shark"
    ],
    r"\bunicorn\b": [
        "mythical", "startup", "billion dollar", "rare", "magical"
    ],
    r"\bdragon\b": [
        "mythical", "power", "fire", "china", "asian"
    ],
    
    # Extended Business & Strategy
    r"strategy|strategic": [
        "planning", "vision", "direction", "goals", "objectives", "roadmap",
        "long term", "competitive", "positioning"
    ],
    r"growth|scale|scaling": [
        "expansion", "increase", "development", "progress", "upward", 
        "acceleration", "hockey stick"
    ],
    r"revenue|sales|income": [
        "money", "earnings", "profit", "top line", "business", "commercial"
    ],
    r"profit|margin|earnings": [
        "money", "bottom line", "return", "gain", "financial performance"
    ],
    r"cost|expense|spend": [
        "money", "budget", "investment", "outflow", "overhead"
    ],
    r"roi|return": [
        "investment", "profit", "gain", "performance", "value"
    ],
    r"market|markets": [
        "industry", "sector", "customers", "demand", "competition", "share"
    ],
    r"customer|client|consumer": [
        "user", "buyer", "account", "relationship", "service"
    ],
    r"stakeholder": [
        "investor", "partner", "interested party", "shareholder", "owner"
    ],
    r"project|initiative": [
        "work", "task", "program", "effort", "undertaking"
    ],
    r"meeting|conference": [
        "discussion", "gathering", "session", "workshop", "collaboration"
    ],
    r"report|reporting": [
        "document", "analysis", "summary", "presentation", "data"
    ],
    r"plan|planning": [
        "strategy", "prepare", "organize", "schedule", "roadmap"
    ],
    r"risk|risky": [
        "danger", "uncertainty", "exposure", "threat", "vulnerability"
    ],
    r"opportunity": [
        "chance", "potential", "prospect", "opening", "possibility"
    ],
    r"challenge|problem": [
        "issue", "difficulty", "obstacle", "hurdle", "barrier"
    ],
    r"solution|solve": [
        "answer", "fix", "resolution", "remedy", "approach"
    ],
    r"innovation|innovate": [
        "new", "creative", "invention", "disruption", "breakthrough"
    ],
    r"transform|transformation": [
        "change", "evolve", "modernize", "digitize", "restructure"
    ],
    r"optimize|optimization": [
        "improve", "enhance", "efficiency", "streamline", "better"
    ],
    r"efficiency|efficient": [
        "productivity", "performance", "streamlined", "lean", "optimal"
    ],
    r"agile|scrum|sprint": [
        "methodology", "iterative", "flexible", "fast", "development"
    ],
    r"kanban|board": [
        "workflow", "tasks", "cards", "columns", "progress"
    ],
    r"lean|kaizen": [
        "efficiency", "continuous improvement", "waste reduction", "optimization"
    ],
    r"six sigma|quality": [
        "process", "improvement", "defects", "variation", "control"
    ],
    
    # Extended Finance & Economics
    r"invest|investment": [
        "capital", "funding", "portfolio", "returns", "assets"
    ],
    r"stock|equity|share": [
        "market", "trading", "ownership", "securities", "investment"
    ],
    r"bond|debt|loan": [
        "fixed income", "credit", "borrowing", "lending", "interest"
    ],
    r"asset|assets": [
        "property", "holdings", "resources", "value", "ownership"
    ],
    r"liability|liabilities": [
        "debt", "obligation", "payable", "owed"
    ],
    r"balance|balance sheet": [
        "accounting", "assets", "liabilities", "equity", "financial statement"
    ],
    r"income statement|p&l|profit and loss": [
        "revenue", "expenses", "profit", "financial statement", "earnings"
    ],
    r"cash flow": [
        "money movement", "liquidity", "inflow", "outflow", "operating"
    ],
    r"forecast|projection": [
        "prediction", "estimate", "future", "planning", "outlook"
    ],
    r"budget|budgeting": [
        "planning", "allocation", "spending", "financial plan", "cost"
    ],
    r"tax|taxation": [
        "government", "revenue", "compliance", "deduction", "filing"
    ],
    r"audit|auditing": [
        "review", "examination", "verification", "compliance", "check"
    ],
    r"compliance|regulatory": [
        "rules", "regulations", "legal", "governance", "standards"
    ],
    r"insurance": [
        "coverage", "protection", "risk", "policy", "premium"
    ],
    r"pension|retirement": [
        "savings", "future", "401k", "benefits", "elderly"
    ],
    r"dividend": [
        "payout", "distribution", "income", "shareholder", "returns"
    ],
    r"merger|acquisition|m&a": [
        "deal", "takeover", "combination", "buyout", "consolidation"
    ],
    r"ipo|public offering": [
        "stock market", "listing", "equity", "shares", "capital raise"
    ],
    
    # Extended Technology & Digital
    r"software|application|app": [
        "program", "system", "platform", "tool", "digital"
    ],
    r"hardware": [
        "device", "equipment", "physical", "computer", "machine"
    ],
    r"digital|digitization": [
        "electronic", "online", "virtual", "technology", "modern"
    ],
    r"artificial intelligence|ai\b": [
        "machine learning", "automation", "smart", "intelligent", "algorithm"
    ],
    r"machine learning|ml\b": [
        "ai", "algorithm", "model", "prediction", "training"
    ],
    r"data science|data scientist": [
        "analytics", "statistics", "modeling", "insights", "big data"
    ],
    r"big data": [
        "analytics", "volume", "variety", "velocity", "processing"
    ],
    r"blockchain|crypto": [
        "decentralized", "ledger", "bitcoin", "ethereum", "web3"
    ],
    r"iot|internet of things": [
        "sensors", "connected", "smart devices", "embedded", "edge"
    ],
    r"5g|connectivity": [
        "network", "wireless", "speed", "mobile", "communication"
    ],
    r"vpn|virtual private": [
        "security", "privacy", "encrypted", "remote", "tunnel"
    ],
    r"firewall": [
        "security", "protection", "network", "barrier", "defense"
    ],
    r"malware|virus|ransomware": [
        "threat", "attack", "security", "hacker", "protection"
    ],
    r"backup|restore": [
        "recovery", "copy", "save", "protection", "redundancy"
    ],
    r"deploy|deployment": [
        "release", "launch", "publish", "rollout", "go live"
    ],
    r"devops|cicd": [
        "automation", "pipeline", "continuous", "integration", "delivery"
    ],
    r"container|docker|kubernetes": [
        "deployment", "orchestration", "microservices", "scalable"
    ],
    r"version|git": [
        "control", "history", "changes", "branch", "commit"
    ],
    r"testing|test|qa": [
        "quality", "verification", "validation", "bug", "check"
    ],
    
    # Extended Operations & Supply Chain
    r"supply chain|procurement": [
        "sourcing", "vendors", "suppliers", "purchasing", "materials"
    ],
    r"logistics|distribution": [
        "shipping", "transportation", "delivery", "movement", "fulfillment"
    ],
    r"inventory|stock": [
        "warehouse", "goods", "products", "materials", "storage"
    ],
    r"order|ordering": [
        "purchase", "request", "buy", "transaction", "fulfillment"
    ],
    r"shipment|ship": [
        "delivery", "transport", "package", "freight", "carrier"
    ],
    r"tracking|track": [
        "monitor", "follow", "status", "location", "progress"
    ],
    r"return|returns": [
        "refund", "exchange", "reverse logistics", "send back"
    ],
    r"supplier|vendor": [
        "provider", "partner", "source", "procurement", "purchasing"
    ],
    r"quality control|qc": [
        "inspection", "standards", "testing", "defects", "compliance"
    ],
    
    # Extended HR & People
    r"hire|hiring|recruit": [
        "employment", "talent", "candidate", "job", "onboard"
    ],
    r"employee|staff|worker": [
        "team member", "personnel", "workforce", "human resources"
    ],
    r"training|learning": [
        "development", "education", "skill", "course", "workshop"
    ],
    r"performance review|evaluation": [
        "assessment", "feedback", "appraisal", "rating", "goals"
    ],
    r"salary|compensation|pay": [
        "wage", "earnings", "income", "benefits", "remuneration"
    ],
    r"benefits|perks": [
        "compensation", "healthcare", "vacation", "retirement", "insurance"
    ],
    r"culture|values": [
        "workplace", "environment", "principles", "beliefs", "company"
    ],
    r"diversity|inclusion|dei": [
        "equity", "belonging", "representation", "equality", "fairness"
    ],
    r"remote|work from home|wfh": [
        "telecommute", "distributed", "virtual", "flexible", "hybrid"
    ],
    r"office|workplace": [
        "building", "headquarters", "location", "space", "environment"
    ],
    
    # Extended Marketing & Sales
    r"marketing|market": [
        "promotion", "advertising", "brand", "campaign", "awareness"
    ],
    r"brand|branding": [
        "identity", "logo", "image", "reputation", "recognition"
    ],
    r"campaign|advertising": [
        "marketing", "promotion", "ads", "media", "reach"
    ],
    r"lead|leads": [
        "prospect", "potential customer", "sales pipeline", "opportunity"
    ],
    r"conversion|convert": [
        "sale", "transaction", "action", "signup", "purchase"
    ],
    r"funnel|pipeline": [
        "stages", "process", "leads", "conversion", "journey"
    ],
    r"crm|customer relationship": [
        "sales", "contacts", "accounts", "pipeline", "relationships"
    ],
    r"seo|search engine": [
        "ranking", "traffic", "keywords", "organic", "visibility"
    ],
    r"social media|social": [
        "facebook", "twitter", "linkedin", "instagram", "engagement"
    ],
    r"content|blog": [
        "article", "post", "writing", "media", "marketing"
    ],
    r"email marketing|newsletter": [
        "campaign", "subscribers", "open rate", "click", "automation"
    ],
    r"webinar|online event": [
        "presentation", "training", "virtual", "video", "audience"
    ],
    
    # Extended ESG & Sustainability
    r"carbon|emission|co2": [
        "climate", "pollution", "footprint", "greenhouse", "environment"
    ],
    r"renewable|clean energy": [
        "solar", "wind", "sustainable", "green", "alternative"
    ],
    r"electric vehicle|ev|tesla": [
        "car", "transportation", "battery", "sustainable", "green"
    ],
    r"waste|garbage|trash": [
        "disposal", "landfill", "reduce", "recycle", "environment"
    ],
    r"plastic|packaging": [
        "material", "waste", "environment", "sustainable", "container"
    ],
    r"organic|natural": [
        "food", "agriculture", "healthy", "chemical-free", "sustainable"
    ],
    r"fair trade|ethical": [
        "social", "responsible", "sustainable", "workers", "supply chain"
    ],
    r"diversity|equity|inclusion": [
        "social", "fairness", "representation", "belonging", "esg"
    ],
    r"governance|board": [
        "oversight", "management", "directors", "corporate", "compliance"
    ],
    r"transparency|disclosure": [
        "reporting", "honesty", "open", "accountability", "trust"
    ],
    
    # Extended Healthcare & Medical
    r"health|healthy|wellness": [
        "medical", "fitness", "wellbeing", "care", "lifestyle"
    ],
    r"hospital|clinic|medical": [
        "healthcare", "doctor", "patient", "treatment", "care"
    ],
    r"doctor|physician|nurse": [
        "medical", "healthcare", "professional", "treatment", "patient"
    ],
    r"medicine|drug|pharmaceutical": [
        "treatment", "prescription", "pill", "healthcare", "therapy"
    ],
    r"vaccine|vaccination": [
        "immunization", "protection", "healthcare", "prevention", "shot"
    ],
    r"dna|genetic|genome": [
        "biology", "science", "hereditary", "research", "biotechnology"
    ],
    r"heart|cardiac": [
        "health", "love", "vital", "cardiovascular", "organ"
    ],
    r"brain|mental": [
        "mind", "cognitive", "thinking", "psychology", "neurology"
    ],
    r"fitness|exercise|gym": [
        "health", "workout", "training", "physical", "active"
    ],
    r"diet|nutrition|food": [
        "health", "eating", "calories", "meal", "wellness"
    ],
    
    # Extended Education & Learning
    r"education|school|university": [
        "learning", "academic", "college", "degree", "study"
    ],
    r"student|learner": [
        "education", "school", "study", "academic", "pupil"
    ],
    r"teacher|professor|instructor": [
        "education", "teaching", "faculty", "mentor", "trainer"
    ],
    r"course|class|lesson": [
        "education", "learning", "training", "module", "curriculum"
    ],
    r"certificate|diploma|degree": [
        "qualification", "credential", "achievement", "education", "graduation"
    ],
    r"exam|test|assessment": [
        "evaluation", "quiz", "score", "grade", "measure"
    ],
    r"library|research": [
        "books", "knowledge", "study", "academic", "information"
    ],
    r"graduate|graduation": [
        "completion", "degree", "achievement", "ceremony", "education"
    ],
    
    # Extended Travel & Transportation
    r"airplane|flight|aviation": [
        "travel", "airport", "flying", "airline", "jet"
    ],
    r"car|automobile|vehicle": [
        "transportation", "driving", "road", "motor", "travel"
    ],
    r"train|railway|metro": [
        "transportation", "rail", "public transit", "commute", "station"
    ],
    r"bus|transit": [
        "transportation", "public", "commute", "route", "passenger"
    ],
    r"ship|boat|cruise": [
        "marine", "ocean", "water", "transportation", "vessel"
    ],
    r"bicycle|bike|cycling": [
        "transportation", "exercise", "green", "pedal", "two wheels"
    ],
    r"taxi|uber|rideshare": [
        "transportation", "ride", "car", "driver", "travel"
    ],
    r"hotel|accommodation|lodging": [
        "travel", "stay", "room", "hospitality", "booking"
    ],
    r"passport|visa|travel document": [
        "travel", "international", "border", "identity", "immigration"
    ],
    r"luggage|suitcase|baggage": [
        "travel", "packing", "trip", "carry", "journey"
    ],
    
    # Visual & Design Elements
    r"circle|round|oval": [
        "shape", "geometric", "curved", "ring", "loop"
    ],
    r"square|rectangle|box": [
        "shape", "geometric", "container", "block", "frame"
    ],
    r"triangle|pyramid": [
        "shape", "geometric", "three", "pointed", "hierarchy"
    ],
    r"star|sparkle": [
        "rating", "favorite", "special", "highlight", "featured"
    ],
    r"heart|love": [
        "favorite", "like", "emotion", "care", "affection"
    ],
    r"home|house": [
        "residence", "property", "dwelling", "homepage", "start"
    ],
    r"pin|marker|location": [
        "map", "place", "position", "point", "destination"
    ],
    r"cursor|pointer|click": [
        "mouse", "select", "interaction", "ui", "interface"
    ],
    r"thumb|thumbs": [
        "like", "dislike", "rating", "approval", "feedback"
    ],
    r"emoji|emoticon|smiley": [
        "face", "expression", "emotion", "reaction", "feeling"
    ],
    
    # Numbers & Quantities
    r"\bone\b|single|solo": [
        "individual", "alone", "unique", "first", "1"
    ],
    r"\btwo\b|dual|pair|double": [
        "couple", "both", "second", "2", "twin"
    ],
    r"\bthree\b|triple|trio": [
        "third", "3", "multiple", "several"
    ],
    r"\bfour\b|quad": [
        "fourth", "4", "multiple", "quarterly"
    ],
    r"\bfive\b|penta": [
        "fifth", "5", "hand"
    ],
    r"zero|empty|none": [
        "nothing", "null", "blank", "0", "clear"
    ],
    r"half|50%|partial": [
        "portion", "incomplete", "middle", "semi"
    ],
    r"full|100%|complete": [
        "entire", "whole", "total", "all", "maximum"
    ],
    r"many|multiple|several": [
        "numerous", "various", "collection", "group", "batch"
    ],
    r"infinite|unlimited|endless": [
        "forever", "continuous", "perpetual", "boundless"
    ],
}

# Category-based enrichments
CATEGORY_ENRICHMENTS = {
    "Business": ["corporate", "enterprise", "professional", "commercial", "company", "organization", "work"],
    "Finance": ["money", "financial", "banking", "monetary", "economic", "investment", "capital", "currency"],
    "Technology": ["tech", "digital", "software", "IT", "computing", "electronic", "innovation", "system"],
    "Communication": ["messaging", "contact", "outreach", "connect", "talk", "dialogue", "correspondence"],
    "Security": ["protection", "safety", "secure", "privacy", "defense", "guard", "access control"],
    "Operations": ["process", "workflow", "logistics", "management", "execution", "production"],
    "Document": ["file", "paper", "content", "record", "report", "attachment", "text"],
    "E-commerce": ["shopping", "retail", "purchase", "commerce", "store", "buy", "sell", "transaction"],
    "ESG": ["sustainability", "environmental", "social", "governance", "green", "climate", "responsible"],
    "Charts": ["data", "analytics", "visualization", "metrics", "graph", "statistics", "reporting"],
    "Map": ["location", "geography", "place", "navigation", "direction", "position", "destination"],
    "Nature": ["environment", "natural", "organic", "ecology", "green", "outdoor", "wildlife"],
    "Media": ["video", "audio", "multimedia", "streaming", "entertainment", "content", "broadcast"],
    "People": ["human", "person", "user", "individual", "team", "group", "workforce"],
    "Health": ["medical", "wellness", "healthcare", "fitness", "care", "treatment"],
    "Arrows": ["direction", "navigation", "movement", "flow", "pointer", "indicator"],
    "Devices": ["hardware", "gadget", "equipment", "electronic", "machine", "tool"],
    "Design": ["creative", "visual", "graphic", "artistic", "layout", "aesthetic"],
    "Brand": ["logo", "identity", "company", "trademark", "recognition"],
    "Weather": ["climate", "temperature", "forecast", "atmospheric", "conditions"],
    "Buildings": ["architecture", "structure", "construction", "real estate", "property"],
    "Vehicles": ["transportation", "travel", "automotive", "mobility", "journey"],
    "Food": ["meal", "dining", "cuisine", "nutrition", "restaurant", "eating"],
    "Games": ["entertainment", "play", "fun", "recreation", "leisure"],
    "Sport": ["athletic", "fitness", "exercise", "competition", "active"],
}


def enrich_icon_tags(icon: dict[str, Any]) -> dict[str, Any]:
    """Add enriched tags to an icon."""
    name = icon.get("name", "").lower()
    icon_id = icon.get("id", "").lower()
    category = icon.get("category", "")
    existing_tags = set(t.lower() for t in icon.get("tags", []))
    
    new_tags = set()
    
    # Apply semantic enrichments based on name and id
    text_to_check = f"{name} {icon_id}"
    for pattern, enrichments in SEMANTIC_ENRICHMENTS.items():
        if re.search(pattern, text_to_check, re.IGNORECASE):
            new_tags.update(enrichments)
    
    # Apply category enrichments
    if category in CATEGORY_ENRICHMENTS:
        new_tags.update(CATEGORY_ENRICHMENTS[category])
    
    # Combine and deduplicate
    all_tags = existing_tags | new_tags
    
    # Update icon
    icon["tags"] = sorted(list(all_tags))
    
    # Rebuild searchable text
    searchable_parts = [
        icon.get("id", ""),
        icon.get("name", ""),
    ] + icon["tags"]
    icon["searchable"] = " ".join(searchable_parts).lower()
    
    return icon


def main():
    parser = argparse.ArgumentParser(description="Enrich icon tags for better searchability")
    parser.add_argument("--preview", "-p", action="store_true", 
                        help="Preview changes without saving")
    parser.add_argument("--source", "-s", type=str, default=None,
                        help="Only enrich icons from specific source (tabler, lucide, etc)")
    parser.add_argument("--sample", type=int, default=0,
                        help="Show sample of enriched icons")
    args = parser.parse_args()
    
    print("Loading icon index...")
    with open(INDEX_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)
    
    total_icons = len(data["icons"])
    print(f"Total icons: {total_icons}")
    
    # Track statistics
    tags_added = 0
    icons_enriched = 0
    
    enriched_samples = []
    
    for icon in data["icons"]:
        # Filter by source if specified
        source_id = icon.get("source_id", icon.get("id", ""))
        if args.source and not source_id.startswith(args.source):
            continue
        
        original_tag_count = len(icon.get("tags", []))
        enrich_icon_tags(icon)
        new_tag_count = len(icon.get("tags", []))
        
        added = new_tag_count - original_tag_count
        if added > 0:
            tags_added += added
            icons_enriched += 1
            
            if len(enriched_samples) < 10:
                enriched_samples.append({
                    "source_id": source_id,
                    "name": icon["name"],
                    "tags_before": original_tag_count,
                    "tags_after": new_tag_count,
                    "sample_tags": icon["tags"][:15],
                })
    
    print(f"\nEnrichment statistics:")
    print(f"  Icons enriched: {icons_enriched}")
    print(f"  Total tags added: {tags_added}")
    print(f"  Average tags added per icon: {tags_added / icons_enriched:.1f}" if icons_enriched > 0 else "")
    
    # Show samples
    if args.sample > 0 or enriched_samples:
        print(f"\nSample enriched icons:")
        for sample in enriched_samples[:args.sample or 5]:
            print(f"  {sample['source_id']}: {sample['name']}")
            print(f"    Tags: {sample['tags_before']} -> {sample['tags_after']}")
            print(f"    Sample: {', '.join(sample['sample_tags'][:10])}")
    
    if args.preview:
        print("\n[Preview mode - no changes saved]")
    else:
        # Save enriched index
        with open(INDEX_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        print(f"\nSaved enriched index to: {INDEX_FILE}")


if __name__ == "__main__":
    main()
