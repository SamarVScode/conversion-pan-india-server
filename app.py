import sys
import logging
from pathlib import Path
from contextlib import asynccontextmanager
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] [%(name)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[logging.StreamHandler(sys.stdout)],
    force=True,
)

SERVER_DIR = Path(__file__).resolve().parent
if str(SERVER_DIR) not in sys.path:
    sys.path.insert(0, str(SERVER_DIR))

from core.jobs import recover_jobs_from_disk
from api.routes import router as api_router
from api.ui import router as ui_router

log = logging.getLogger("conversion_report_server.app")


@asynccontextmanager
async def lifespan(app: FastAPI):
    log.info("Booting Conversion Report Stream Server: recovering job state from disk...")
    recover_jobs_from_disk()
    yield
    log.info("Shutting down Conversion Report Stream Server.")


def create_app() -> FastAPI:
    app = FastAPI(
        title="Conversion Report Stream Server",
        description="Standalone low-memory streaming FastAPI server for Conversion Report (DC_View & Agent_View)",
        version="1.0.0",
        lifespan=lifespan,
    )

    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )

    app.include_router(api_router)
    app.include_router(ui_router)

    return app
