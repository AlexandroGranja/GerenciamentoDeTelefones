from fastapi import FastAPI

from .api.v1.endpoints.auth import router as auth_router


app = FastAPI(title="Sistema de Chamados - API")

# Endpoint esperado pelo fluxo de SSO do React.
app.include_router(auth_router)

