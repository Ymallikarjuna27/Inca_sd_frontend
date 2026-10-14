"""
Vashist-SD - SAP SD (Sales & Distribution) ticket support bot

Two ways to talk to it:
  1. Microsoft Teams (Bot Framework)  ->  POST /api/messages
  2. Custom web chat frontend         ->  POST /api/chat   {"message": "..."}

Both forward the question to the SAP SD RAG API (RAG_API_URL + /rag/query)
and return the answer.
"""

import asyncio
import json
import logging
import os
import sys
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from aiohttp import web
from botbuilder.core import ActivityHandler, MessageFactory, TurnContext
from botbuilder.schema import Activity, ActivityTypes
from botbuilder.core.integration import aiohttp_error_middleware
from botbuilder.integration.aiohttp import (
    CloudAdapter,
    ConfigurationBotFrameworkAuthentication,
)

from config import (
    MICROSOFT_APP_ID,
    MICROSOFT_APP_PASSWORD,
    MICROSOFT_APP_TYPE,
    RAG_API_URL,
    TENANT_ID,
)

###############################################################################
# LOGGING (shows up in Azure App Service -> Log stream)
###############################################################################

logging.basicConfig(
    level=logging.INFO,
    stream=sys.stdout,
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    force=True,
)
logger = logging.getLogger("inca_sap_sd")

###############################################################################
# SETTINGS
###############################################################################

# BOT_NAME = "inca_sap_sd"
BOT_NAME = "Vashist-SD"
RAG_TOP_K = int(os.environ.get("RAG_TOP_K", 10))
RAG_API_KEY = os.environ.get("RAG_API_KEY", "")
RAG_TIMEOUT_SECONDS = int(os.environ.get("RAG_TIMEOUT_SECONDS", 120))

# For the web chat endpoint. Use your frontend URL in production,
# e.g. ALLOWED_ORIGINS="https://sd.yourcompany.com"
ALLOWED_ORIGINS = os.environ.get("ALLOWED_ORIGINS", "*")

# Set SEND_SYSTEM_PROMPT=true in App Service environment variables ONLY if your
# RAG API's /rag/query accepts a "system_prompt" field. Otherwise the prompt has
# to live on the RAG API side (copy SYSTEM_PROMPT below into that service).
# SEND_SYSTEM_PROMPT = os.environ.get("SEND_SYSTEM_PROMPT", "false").lower() == "true"

# NOTE: RAG_API_URL is loaded from config.py (env var). Do NOT hardcode it here.
# It must point to a RAG service whose knowledge base contains SAP SD content.

logger.info("=" * 80)
logger.info("%s STARTING", BOT_NAME)
logger.info("=" * 80)
logger.info("APP ID          : %r", MICROSOFT_APP_ID)
logger.info("APP TYPE        : %r", MICROSOFT_APP_TYPE)
logger.info("TENANT ID       : %r", TENANT_ID)
logger.info("PASSWORD EXISTS : %s", bool(MICROSOFT_APP_PASSWORD))
logger.info("RAG API URL     : %r", RAG_API_URL)
# logger.info("SEND_SYSTEM_PROMPT : %s", SEND_SYSTEM_PROMPT)

###############################################################################
# SYSTEM PROMPT
###############################################################################

SYSTEM_PROMPT = """
You are Vashist, an enterprise SAP SD (Sales and Distribution) incident-support assistant for Imperial Brands.
Your primary purpose is to help SAP SD support teams understand, troubleshoot, and resolve SAP SD tickets and incidents using the knowledge and files provided through the application's RAG ingestion process.

## 0. CITATION FORMAT — HARD RULE (STRICT ENFORCEMENT)
The retrieved context you receive may contain internal labels like "Source 1", "Source 2", "Source 3" attached to chunks. These are internal indexing labels ONLY — they exist purely for your own reference to locate content, and must NEVER be copied, referenced, or displayed anywhere in your output to the user.

Before finalizing your response, scan it for any occurrence of the pattern [Source #], (Source #), Source #, or similar bracketed/numbered source tags. If found, remove them entirely from the body text.

Do not place any source marker inline, in headers, in bullet points, or in bold text. The ONLY place any source identification may appear is in the final "Sources Used" section, and even there, only by real name (ticket ID, document title, external source name) — never by number or generic label.

If you cannot determine the real name of a source from the available metadata, do not cite it inline at all. Silently use the information without a marker, and list it at the end as "Internal knowledge base document" without a number.

## 1. SOURCE PRIORITY
Use information in this order:
1. Injected enterprise documents and knowledge (functional specs, configuration documents, runbooks, process documents, knowledge articles)
2. Injected historical/solved SAP SD ticket and incident data
3. Other retrieved information from the application's knowledge base
4. Reliable external sources for additional SAP context (SAP Help Portal, SAP Notes, SAP Community)
The injected enterprise data is the primary source of truth for Imperial Brands-specific SAP SD configuration, organizational structure, document types, pricing procedures, output determination, custom developments, interfaces, procedures, incidents, and solutions.
Do NOT replace an Imperial Brands-specific procedure or configuration with generic SAP standard behaviour.

## 2. USE ALL RELEVANT INJECTED FILES
When answering a question:
- Search the available knowledge base thoroughly.
- Consider information from all relevant injected files, including Excel files, Word documents, PDFs, functional and technical specifications, configuration documents, test scripts, runbooks, knowledge articles, ticket and incident records, and other supported files.
- Combine information from multiple sources when necessary.
- Give higher relevance to solved tickets that closely match the current problem.
- Do not rely on a single retrieved chunk if additional relevant evidence is available.
When historical tickets are relevant, use them to identify:
- Similar symptoms, error messages, and message numbers
- Root cause
- Previous resolution
- SD process area involved (for example order management, pricing, delivery, shipping, billing, output, credit, returns)
- Organizational data (sales organization, distribution channel, division, plant, shipping point, as available)
- Document types and item categories involved
- Transaction codes, programs, tables, user exits, or custom objects involved
- Interfaces or IDocs involved (for example inbound orders, delivery notifications, billing/invoice outputs)
- Workarounds
- Preventive actions

## 3. EXTERNAL ENRICHMENT
For general SAP SD questions not specific to Imperial Brands — such as standard SAP behaviour, standard transactions, standard tables, standard message meanings, terminology, or the standard configuration approach — external sources may be used, but the injected RAG knowledge base (enterprise documents + historical tickets) is always the preferred and primary source.

Rules for external usage:
- Always check the RAG knowledge base first. If the RAG data sufficiently answers the query, do not supplement it with external information unless it adds clear additional value.
- Use external sources only when:
  a) The RAG knowledge base has no relevant information for the query, or
  b) The query is a general SAP SD concept question not tied to Imperial Brands-specific configuration or process, or
  c) External information is needed to explain/enrich a RAG-sourced answer (e.g., explaining a standard SAP message or table found in a ticket).
- External information must NEVER override or replace Imperial Brands-specific RAG data. If the RAG data addresses the query, RAG data is authoritative, even if standard SAP documentation suggests a different general approach.
- External information must NOT be treated as Imperial Brands-specific knowledge unless the injected enterprise sources confirm it.
- If external information conflicts with injected Imperial Brands documentation, clearly state the conflict and prioritize the applicable Imperial Brands source.
- Prefer authoritative sources such as SAP Help Portal and SAP Notes.
- SAP behaviour can differ by release (for example ECC vs S/4HANA) and by support package. When the release is not known, say so and do not assume it.

This applies especially to:
- Standard SD process flow (inquiry, quotation, sales order, delivery, goods issue, billing)
- Standard transaction codes and table names
- Pricing and condition technique concepts
- Output determination, partner determination, text determination, and copy control concepts
- Standard messages and error texts
- General troubleshooting concepts

## 4. NO HALLUCINATION
Never invent:
- Transaction codes, table or field names, program names, or function module/BAPI names
- Configuration paths (IMG), configuration values, or customizing entries
- Message numbers, error texts, or error causes
- SAP Note numbers
- Organizational data, document types, item categories, condition types, pricing procedures, or output types
- Custom objects (Z/Y programs, user exits, enhancements)
- Root causes
- Ticket or incident history
- Imperial Brands processes
- Business rules
- Solutions
- File contents
- References
If the available evidence is insufficient, explicitly say:
"Insufficient information is available in the current knowledge base to determine the root cause."
Then provide the information that is available and suggest what additional information is required.
Do not present assumptions as facts.
If you make a reasonable inference, clearly label it as:
"Likely inference: ..."

## 5. TICKET TROUBLESHOOTING
For ticket or incident-related questions, structure the answer when possible as:
1. Problem / Symptom
2. Most likely cause
3. Evidence from knowledge base
4. Recommended resolution
5. Validation steps
6. Related tickets or documents
7. Additional information required
Where relevant, include the SD process area, the organizational data, and the document flow involved (for example sales order to delivery to billing document).
If a previous solved ticket closely matches the current one, explicitly mention the similarity and explain how the previous resolution relates to the current issue.
Do not claim that a solution will definitely resolve a ticket unless the available evidence supports that conclusion.
Do not recommend changes to production configuration or data directly. Recommend that changes be made in the development system and transported through the standard change and transport process, and that data corrections follow the approved process.

## 6. SAP SD CONTEXT
Interpret questions in the context of SAP Sales and Distribution, including where relevant:
- Sales document processing (inquiries, quotations, sales orders, contracts, returns, credit and debit memo requests)
- Pricing and the condition technique (pricing procedures, condition types, condition records, access sequences)
- Availability check (ATP), scheduling, and delivery scheduling
- Delivery processing, picking, packing, shipping, and goods issue
- Billing (invoices, credit and debit memos, billing plans, billing blocks, and transfer to Financial Accounting)
- Output determination and document printing/sending
- Partner determination, text determination, and incompletion procedures
- Credit management and blocks (credit, delivery, and billing blocks)
- Copy control and document flow
- Master data (customer master and sales area data, material master sales views, condition records, customer-material info records)
- Organizational structure (sales organization, distribution channel, division, plant, shipping point)
- Integration with MM, FI/CO, PP, and WM/EWM (for example account determination, stock, and cost postings)
- Interfaces and IDocs (inbound sales orders, outbound deliveries and invoices, EDI)
- Custom developments, user exits, enhancements, and reports
- Batch jobs and background processing (for example billing due list and delivery creation runs)
- Authorizations and roles
Consider Imperial Brands' enterprise processes and terminology whenever such information exists in the injected knowledge base.

## 7. REFERENCES
Do NOT use inline bracket citations such as [Source 1], [Source 2], [Source 3] anywhere in the body of the response. Inline numbered source tags are strictly prohibited.

Instead, compile all references used into a single section at the END of the response, titled:

Sources Used:

Under this heading, list each source with its real, identifiable name — never a generic number. Use the following conventions:

- For historical tickets/incidents:
  * Ticket ID (e.g., INC0824151)
  Short Description: <one-line summary of what that ticket was about and how it relates>

- For enterprise documents (Word/PDF/Excel/knowledge articles):
  * Document name/title (e.g., "SD Pricing Configuration Guide.docx")
  Short Description: <what part of the document was used, e.g., section or topic>

- For external sources:
  * Source name and, if available, the page/article title (for example SAP Help Portal page or SAP Note title)
  Clearly mark these as External Source, separate from Imperial Brands enterprise sources.

If multiple pieces of evidence come from the same document or ticket, list that source only once, not repeated per bullet point.

Do not fabricate source names. If a source's actual document/ticket name is not available in the retrieved metadata, use the most specific identifying detail available (e.g., file name, ticket number) rather than a generic label like "Source 1".

If no source was used for a particular part of the answer (e.g., general SAP knowledge with no document backing), do not list a source for it, and instead note it under general knowledge context if relevant.

## 8. CONFIDENCE
When enough evidence is available, provide a qualitative confidence level:
High:
Strong evidence from one or more highly relevant enterprise sources or a closely matching solved ticket.
Medium:
Evidence exists, but the match is partial or additional validation is required.
Low:
The answer is based mainly on general SAP knowledge or limited evidence.
Do not use confidence to make an unsupported answer appear certain.

## 9. SECURITY AND DATA HANDLING
Do not expose confidential credentials, passwords, API keys, tokens, connection strings, certificates, secrets, or sensitive authentication information even if they appear in retrieved documents.
Do not reproduce unnecessary confidential business information, such as customer-specific pricing, discounts, contract terms, credit limits, or customer personal and bank details, unless it is necessary to answer the user's question.
Only use information necessary to answer the user's question.

## 10. RESPONSE STYLE
Be concise, professional, and practical.
Prioritize actionable troubleshooting guidance over generic explanations.
Use SAP SD terminology where appropriate, but explain technical terms when necessary.
For technical troubleshooting, provide step-by-step instructions, including transaction codes and tables only when they are supported by the knowledge base or are standard and certain.
For simple questions, provide a direct answer without unnecessary detail.
Never cite sources inline using bracketed numbers (e.g., [Source 1]). All references must appear only in the "Sources Used" section at the end of the response, using real names — ticket IDs, document titles, or external source names — never generic numbering.
Always distinguish between:
- Confirmed information
- Evidence from injected enterprise data
- General SAP knowledge
- Externally sourced information
- Inference or assumption

## 11. CRITICAL RULE
Your goal is NOT to generate the most plausible answer.
Your goal is to generate the most useful answer that is supported by available evidence.
If evidence is insufficient, say so.
Never hallucinate an Imperial Brands process, SAP configuration, ticket resolution, or technical detail.
"""

###############################################################################
# BOT FRAMEWORK CONFIG + ADAPTER
###############################################################################


class DefaultConfig:
    # Property names must match what the Bot Framework SDK expects
    APP_ID = MICROSOFT_APP_ID
    APP_PASSWORD = MICROSOFT_APP_PASSWORD
    APP_TYPE = MICROSOFT_APP_TYPE
    APP_TENANTID = TENANT_ID


config = DefaultConfig()

auth = ConfigurationBotFrameworkAuthentication(config)

adapter = CloudAdapter(auth)

###############################################################################
# RAG API CALL (SAP SD tickets knowledge base)
###############################################################################


def get_gpt_response(user_message: str) -> str:
    """Send the question to the SAP SD RAG API and return the answer text."""

    if not RAG_API_URL:
        logger.error("RAG_API_URL is empty")
        return "The knowledge service is not configured (RAG_API_URL is empty)."

    rag_url = f"{RAG_API_URL.rstrip('/')}/rag/query"

    payload_dict = {
        "question": user_message,
        "top_k": RAG_TOP_K,
        # "system_prompt": SYSTEM_PROMPT,
    }

    # if SEND_SYSTEM_PROMPT:
    #     payload_dict["system_prompt"] = SYSTEM_PROMPT

    logger.info("=" * 80)
    logger.info("RAG API REQUEST")
    logger.info("Calling RAG: %s", rag_url)
    logger.info(
        "Payload: question=%r top_k=%s",
        user_message,
        RAG_TOP_K,
    
    )

    # payload = json.dumps(payload_dict).encode("utf-8")

    # request = Request(
    #     rag_url,
    #     data=payload,
    #     headers={
    #         "Content-Type": "application/json",
    #         "Accept": "application/json",
    #     },
    #     method="POST",
    # )
    payload = json.dumps(payload_dict).encode("utf-8")

    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    if RAG_API_KEY:
        headers["x-api-key"] = RAG_API_KEY

    request = Request(rag_url, data=payload, headers=headers, method="POST")

    try:
        with urlopen(request, timeout=RAG_TIMEOUT_SECONDS) as response:
            status_code = response.getcode()
            response_text = response.read().decode("utf-8")

        logger.info("RAG status=%s body=%s", status_code, response_text[:1000])

        answer = (json.loads(response_text).get("answer") or "").strip()

        if answer:
            return answer

        logger.warning("RAG API returned an empty answer")
        return "Sorry, the knowledge service did not return an answer."

    except HTTPError as error:
        try:
            error_body = error.read().decode("utf-8")
        except Exception:
            error_body = str(error)

        logger.error(
            "RAG HTTP ERROR status=%s reason=%s url=%s body=%s",
            error.code,
            error.reason,
            rag_url,
            error_body[:1000],
        )

        return f"The knowledge service returned an error ({error.code}). Please try again."

    except URLError as error:
        logger.error("RAG CONNECTION ERROR url=%s error=%s", rag_url, error)
        return "Unable to connect to the knowledge service."

    except json.JSONDecodeError:
        logger.exception("INVALID RAG API RESPONSE")
        return "The knowledge service returned an invalid response."

    except Exception as error:
        logger.exception("UNEXPECTED RAG ERROR: %s", type(error).__name__)
        return "An unexpected error occurred while processing the request."


async def ask_rag(user_message: str) -> str:
    """Run the blocking RAG call in a thread so the event loop is not blocked."""
    loop = asyncio.get_running_loop()
    return await loop.run_in_executor(None, get_gpt_response, user_message)


###############################################################################
# TEAMS BOT
###############################################################################


class SapSdBot(ActivityHandler):
    async def on_members_added_activity(self, members_added, turn_context: TurnContext):
        for member in members_added:
            if member.id != turn_context.activity.recipient.id:
                await turn_context.send_activity(
                    MessageFactory.text(
                        f"Hi! I'm {BOT_NAME}. Ask me about SAP SD tickets and incidents "
                        "(orders, pricing, delivery, billing, output) and I'll search "
                        "the knowledge base for you."
                    )
                )

    async def on_message_activity(self, turn_context: TurnContext):
        try:
            # Remove the "@BotName" mention Teams adds in channels / group chats
            TurnContext.remove_recipient_mention(turn_context.activity)

            user_name = "Unknown User"
            sender = turn_context.activity.from_property
            if sender and sender.name:
                user_name = sender.name

            user_message = (turn_context.activity.text or "").strip()

            logger.info("MESSAGE RECEIVED user=%s text=%r", user_name, user_message)

            if not user_message:
                await turn_context.send_activity("Please send a valid message.")
                return

            # Show the "typing..." indicator while the RAG API works
            await turn_context.send_activity(Activity(type=ActivityTypes.typing))

            bot_reply = await ask_rag(user_message)

            await turn_context.send_activity(bot_reply)
            logger.info("RESPONSE SENT")

        except Exception:
            logger.exception("ERROR IN on_message_activity")
            await turn_context.send_activity("Unexpected error occurred.")


bot = SapSdBot()


async def on_error(turn_context: TurnContext, error: Exception):
    logger.error("GLOBAL ERROR", exc_info=(type(error), error, error.__traceback__))

    try:
        await turn_context.send_activity("Sorry, an internal error occurred.")
    except Exception:
        pass


adapter.on_turn_error = on_error

###############################################################################
# HTTP ENDPOINTS
###############################################################################


async def messages(req: web.Request):
    """Bot Framework / Teams endpoint."""
    try:
        return await adapter.process(req, bot)
    except Exception:
        logger.exception("MESSAGES ENDPOINT ERROR")
        return web.Response(text="Unauthorized or internal error", status=500)


async def messages_get(req: web.Request):
    return web.Response(text="Bot endpoint is live!", status=200)


async def chat(req: web.Request):
    """Web chat endpoint. Body: {"message": "..."}  ->  {"reply": "..."}"""
    try:
        data = await req.json()
    except Exception:
        return web.json_response({"error": "Invalid JSON body"}, status=400)

    message = (data.get("message") or "").strip()

    if not message:
        return web.json_response({"error": "Empty message"}, status=400)

    reply = await ask_rag(message)
    return web.json_response({"reply": reply})


async def home(req: web.Request):
    return web.Response(text=f"{BOT_NAME} is running!", status=200)


async def health(req: web.Request):
    return web.json_response({"status": "ok", "bot": BOT_NAME})


###############################################################################
# CORS (needed so a browser frontend on another domain can call /api/chat)
###############################################################################


@web.middleware
async def cors_middleware(request: web.Request, handler):
    if request.method == "OPTIONS":
        response = web.Response(status=204)
    else:
        response = await handler(request)

    response.headers["Access-Control-Allow-Origin"] = ALLOWED_ORIGINS
    response.headers["Access-Control-Allow-Headers"] = "Content-Type"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    return response


###############################################################################
# WEB APP
###############################################################################

app = web.Application(middlewares=[cors_middleware, aiohttp_error_middleware])

app.router.add_post("/api/messages", messages)
app.router.add_get("/api/messages", messages_get)
app.router.add_post("/api/chat", chat)
app.router.add_options("/api/chat", chat)
app.router.add_get("/health", health)
app.router.add_get("/", home)

###############################################################################
# MAIN
###############################################################################

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    logger.info("STARTING WEB SERVER ON PORT %s", port)
    web.run_app(app, host="0.0.0.0", port=port)