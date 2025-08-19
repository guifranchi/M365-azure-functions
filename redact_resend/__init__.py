import azure.functions as func
import json
async def main(req: func.HttpRequest) -> func.HttpResponse:
    body = {"status":"resent","messageId":"dummy-id"}
    return func.HttpResponse(json.dumps(body),status_code=200,mimetype="application/json")
