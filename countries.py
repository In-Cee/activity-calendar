# ============================================================
#  Country flag emoji lookup - v0.4
# ============================================================

COUNTRY_FLAGS = {
    "Nigeria": "🇳🇬",
    "Ghana": "🇬🇭",
    "Senegal": "🇸🇳",
    "Cote d'Ivoire": "🇨🇮",
    "Côte d'Ivoire": "🇨🇮",
    "Mali": "🇲🇱",
    "Burkina Faso": "🇧🇫",
    "Kenya": "🇰🇪",
    "Tanzania": "🇹🇿",
    "Uganda": "🇺🇬",
    "Rwanda": "🇷🇼",
    "Ethiopia": "🇪🇹",
    "South Africa": "🇿🇦",
    "Malawi": "🇲🇼",
    "Zambia": "🇿🇲",
    "Mozambique": "🇲🇿",
    "DR Congo": "🇨🇩",
    "Morocco": "🇲🇦",
    "Egypt": "🇪🇬",
    "Tunisia": "🇹🇳",
    "Canada": "🇨🇦",
    "United States": "🇺🇸",
    "United Kingdom": "🇬🇧",
}

def flag(country: str) -> str:
    """Return the flag emoji for a country, or a globe if unknown."""
    return COUNTRY_FLAGS.get(country, "🌍")
