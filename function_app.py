import azure.functions as func
import logging
import asyncio

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)

@app.route(route="redact_resend")
async def redact_resend(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Python HTTP trigger function processed a request.')

    try:
        # Import your existing function logic
        from redact_resend import main
        
        # Call your existing async function
        return await main(req)
        
    except Exception as e:
        logging.error(f"Error in redact_resend function: {str(e)}")
        return func.HttpResponse(
            "An error occurred while processing the request.",
            status_code=500
        )
