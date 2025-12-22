from fastapi import FastAPI, Request
import requests
import os

app = FastAPI()

# Mensagens por etapa
MENSAGENS = {
    "programado para produção": "Seu projeto foi programado para produção.",
    "em produção": "Seu projeto está em produção.",
    "finalizado": "Seu projeto foi finalizado."
}

@app.post("/clickup")
async def receber_clickup(request: Request):
    data = await request.json()

    for item in data.get("history_items", []):
        if item.get("field") == "status":
            novo_status = item["after"]["status"].lower()

            if novo_status in MENSAGENS:
                ticket_id = buscar_ticket_zendesk(data)
                if ticket_id:
                    enviar_mensagem_zendesk(ticket_id, MENSAGENS[novo_status])

    return {"ok": True}

def buscar_ticket_zendesk(data):
    """
    Aqui você PRECISA ter um campo no ClickUp
    com o ID do ticket do Zendesk.
    """
    for campo in data["task"].get("custom_fields", []):
        if campo["name"].lower() == "zendesk_ticket_id":
            return campo["value"]
    return None

def enviar_mensagem_zendesk(ticket_id, mensagem):
    url = f"https://{os.environ['ZENDESK_DOMAIN']}/api/v2/tickets/{ticket_id}.json"

    payload = {
        "ticket": {
            "comment": {
                "body": mensagem,
                "public": True
            }
        }
    }

    requests.put(
        url,
        json=payload,
        auth=(f"{os.environ['ZENDESK_EMAIL']}/token", os.environ['ZENDESK_API_TOKEN'])
    )
