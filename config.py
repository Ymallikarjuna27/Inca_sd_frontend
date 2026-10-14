import os

from dotenv import load_dotenv

# Loads .env for local development. In Azure, real environment
# variables are used instead and .env does not exist.
load_dotenv()


def _require(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if not value:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return value


# --- Bot / Azure AD (required) ---
MICROSOFT_APP_ID = _require("MICROSOFT_APP_ID")
MICROSOFT_APP_PASSWORD = _require("MICROSOFT_APP_PASSWORD")
MICROSOFT_APP_TYPE = os.environ.get("MICROSOFT_APP_TYPE", "SingleTenant")
TENANT_ID = _require("TENANT_ID")

# --- Azure OpenAI (optional: not used by app.py, since the RAG API calls the model) ---
AZURE_OPENAI_KEY = os.environ.get("AZURE_OPENAI_KEY", "")
AZURE_OPENAI_ENDPOINT = os.environ.get("AZURE_OPENAI_ENDPOINT", "")
AZURE_OPENAI_DEPLOYMENT = os.environ.get("AZURE_OPENAI_DEPLOYMENT", "gpt-4o")
AZURE_OPENAI_API_VERSION = os.environ.get("AZURE_OPENAI_API_VERSION", "2024-12-01-preview")

# RAG_API_URL = _require("RAG_API_URL")
RAG_API_KEY = os.environ.get("RAG_API_KEY", "")

# --- RAG API (required) ---
RAG_API_URL = _require("RAG_API_URL")



