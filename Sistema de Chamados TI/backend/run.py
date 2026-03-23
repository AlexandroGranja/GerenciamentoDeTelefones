import os

import uvicorn


if __name__ == "__main__":
    host = os.environ.get("CHAMADOS_API_HOST", "0.0.0.0")
    port = int(os.environ.get("CHAMADOS_API_PORT", "8000"))
    uvicorn.run("app.main:app", host=host, port=port, reload=False)

