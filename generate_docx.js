"use strict";
const DOCX_PATH = "C:/Users/ptawade/AppData/Roaming/npm/node_modules/docx";
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageNumber, Header, Footer
} = require(DOCX_PATH);
const fs = require("fs");
const path = require("path");

const OUT = "C:/Users/ptawade/Desktop/SNOP Product/snop/SNOP_Backend_Code.docx";

// ── colours / fonts ──────────────────────────────────────────────────────────
const DARK_BLUE  = "1F3864";
const DARK_GREY  = "404040";
const MED_GREY   = "595959";
const LIGHT_BG   = "F2F2F2";
const CODE_FONT  = "Courier New";
const BODY_FONT  = "Arial";

// ── helpers ──────────────────────────────────────────────────────────────────

function titlePara(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 240, after: 80 },
    children: [new TextRun({ text, font: BODY_FONT, size: 72, bold: true, color: "000000" })],
  });
}

function subtitlePara(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 480 },
    children: [new TextRun({ text, font: BODY_FONT, size: 28, color: MED_GREY })],
  });
}

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 480, after: 160 },
    children: [new TextRun({ text, font: BODY_FONT, size: 36, bold: true, color: DARK_BLUE })],
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 80 },
    shading: { fill: LIGHT_BG, type: ShadingType.CLEAR },
    children: [new TextRun({ text, font: CODE_FONT, size: 22, bold: true, color: DARK_GREY })],
  });
}

function heading3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 200, after: 80 },
    children: [new TextRun({ text, font: BODY_FONT, size: 24, bold: true, color: DARK_GREY })],
  });
}

/** Render source code: one Paragraph per line, preserving indentation. */
function codeBlock(source) {
  const lines = source.split("\n");
  // Remove a trailing blank line artifact if present
  if (lines.length > 0 && lines[lines.length - 1].trim() === "") lines.pop();
  return lines.map(line => new Paragraph({
    spacing: { before: 0, after: 0, line: 200, lineRule: "exact" },
    children: [new TextRun({
      text: line === "" ? " " : line,  // keep blank lines visible
      font: CODE_FONT,
      size: 16,  // 8pt
      color: "1A1A1A",
    })],
  }));
}

function spacer(before = 80) {
  return new Paragraph({ spacing: { before, after: 0 }, children: [new TextRun("")] });
}

// ── table helper ─────────────────────────────────────────────────────────────
const BORDER = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const BORDERS = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER };

function makeTable(headers, rows, colWidths) {
  const totalW = colWidths.reduce((a,b)=>a+b, 0);
  const makeCell = (text, isHeader, w) => new TableCell({
    borders: BORDERS,
    width: { size: w, type: WidthType.DXA },
    shading: { fill: isHeader ? "D5E8F0" : "FFFFFF", type: ShadingType.CLEAR },
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({
      spacing: { before: 0, after: 0 },
      children: [new TextRun({ text, font: BODY_FONT, size: 18, bold: isHeader, color: "000000" })],
    })],
  });
  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => makeCell(h, true, colWidths[i])),
  });
  const dataRows = rows.map(row => new TableRow({
    children: row.map((cell, i) => makeCell(cell, false, colWidths[i])),
  }));
  return new Table({
    width: { size: totalW, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [headerRow, ...dataRows],
  });
}

// ── all source files ──────────────────────────────────────────────────────────

const MAIN_PY = `"""SNOP Platform — FastAPI Application Entry Point"""
from contextlib import asynccontextmanager
import structlog
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.middleware.gzip import GZipMiddleware

from app.core.config   import settings
from app.core.database import engine, Base
from app.api.routes    import (
    health, dashboard, plants, furnaces,
    feeds, yield_matrix, optimization, scenarios, uploads, regression, sensitivity
)

log = structlog.get_logger()


@asynccontextmanager
async def lifespan(app: FastAPI):
    log.info("SNOP starting", env=settings.ENVIRONMENT)
    async with engine.begin() as conn:
        await conn.run_sync(Base.metadata.create_all)
    log.info("Database schema ready")
    yield
    await engine.dispose()


app = FastAPI(
    title="SNOP Platform API",
    description="Sales & Operations Planning — Multi-Plant Olefin Units",
    version="2.0.0",
    docs_url="/api/docs",
    redoc_url="/api/redoc",
    openapi_url="/api/openapi.json",
    lifespan=lifespan,
)

app.add_middleware(GZipMiddleware, minimum_size=1000)
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.CORS_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(health.router)
app.include_router(dashboard.router,    prefix="/api/dashboard",    tags=["Dashboard"])
app.include_router(plants.router,       prefix="/api/plants",       tags=["Plants"])
app.include_router(furnaces.router,     prefix="/api/furnaces",     tags=["Furnaces"])
app.include_router(feeds.router,        prefix="/api/feeds",        tags=["Feeds"])
app.include_router(yield_matrix.router, prefix="/api/yield-matrix", tags=["Yield Matrix"])
app.include_router(optimization.router, prefix="/api/optimization", tags=["Optimization"])
app.include_router(scenarios.router,    prefix="/api/scenarios",    tags=["Scenarios"])
app.include_router(uploads.router,      prefix="/api/uploads",      tags=["Uploads"])
app.include_router(regression.router,   prefix="/api/regression",   tags=["Regression"])
app.include_router(sensitivity.router,  prefix="/api/sensitivity",  tags=["Sensitivity"])`;

const CONFIG_PY = `from typing import List
from pydantic import field_validator
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=".env", extra="ignore")

    ENVIRONMENT: str = "development"
    SECRET_KEY: str = "dev_secret_key_change_in_production"
    DATABASE_URL: str = "postgresql+asyncpg://snop_admin:snop_secret@localhost:5433/snop_production"
    SYNC_DATABASE_URL: str = "postgresql://snop_admin:snop_secret@localhost:5433/snop_production"
    REDIS_URL: str = "redis://localhost:6379/0"
    CORS_ORIGINS: List[str] = ["http://localhost:5173", "http://localhost:3000"]
    UPLOAD_DIR: str = "uploads"
    MAX_UPLOAD_MB: int = 50
    OPT_TIMEOUT_SEC: int = 600

    @field_validator("CORS_ORIGINS", mode="before")
    @classmethod
    def parse_cors(cls, v):
        if isinstance(v, str):
            v = v.strip()
            if v.startswith("["):
                import json
                return json.loads(v)
            return [i.strip() for i in v.split(",")]
        return v

    @property
    def is_production(self) -> bool:
        return self.ENVIRONMENT == "production"


settings = Settings()`;

const DATABASE_PY = `from typing import AsyncGenerator
from sqlalchemy.ext.asyncio import AsyncSession, async_sessionmaker, create_async_engine
from sqlalchemy.orm import DeclarativeBase
from app.core.config import settings

engine = create_async_engine(
    settings.DATABASE_URL,
    pool_size=20, max_overflow=40, pool_pre_ping=True,
    echo=not settings.is_production,
)

AsyncSessionLocal = async_sessionmaker(
    bind=engine, class_=AsyncSession,
    expire_on_commit=False, autoflush=False,
)


class Base(DeclarativeBase):
    pass


async def get_db() -> AsyncGenerator[AsyncSession, None]:
    async with AsyncSessionLocal() as session:
        try:
            yield session
            await session.commit()
        except Exception:
            await session.rollback()
            raise
        finally:
            await session.close()`;

const CELERY_PY = `from celery import Celery
from app.core.config import settings

celery_app = Celery(
    "snop",
    broker=settings.REDIS_URL,
    backend=settings.REDIS_URL,
    include=["app.services.optimization_tasks"],
)

celery_app.conf.update(
    task_serializer="json",
    result_serializer="json",
    accept_content=["json"],
    timezone="UTC",
    enable_utc=True,
    task_track_started=True,
    task_routes={"app.services.optimization_tasks.*": {"queue": "optimization"}},
    task_time_limit=settings.OPT_TIMEOUT_SEC + 120,
    task_soft_time_limit=settings.OPT_TIMEOUT_SEC,
    worker_prefetch_multiplier=1,
)`;

const MODELS_PY = `import uuid, enum
from datetime import datetime
from typing import Optional, List
from sqlalchemy import (
    String, Float, Integer, Boolean, DateTime, Text, JSON, LargeBinary,
    ForeignKey, Enum as SAEnum, UniqueConstraint, Index,
)
from sqlalchemy.dialects.postgresql import UUID
from sqlalchemy.orm import Mapped, mapped_column, relationship
from sqlalchemy.sql import func
from app.core.database import Base


class PlantStatus(str, enum.Enum):
    ACTIVE = "active"; SHUTDOWN = "shutdown"; PLANNED = "planned"

class FurnaceStatus(str, enum.Enum):
    RUNNING = "running"; DECOKING = "decoking"; STANDBY = "standby"; MAINTENANCE = "maintenance"

class CoilType(str, enum.Enum):
    SHORT_RESIDENCE = "short_residence"; LONG_RESIDENCE = "long_residence"
    MILLISECOND = "millisecond"; ULTRA_SELECTIVE = "ultra_selective"
    # Lummus SRT family
    SRT_I   = "srt_i";   SRT_II  = "srt_ii";  SRT_III = "srt_iii"
    SRT_IV  = "srt_iv";  SRT_V   = "srt_v";   SRT_VI  = "srt_vi";  SRT_VII = "srt_vii"
    # Technip
    GK_6    = "gk_6";    SMK     = "smk"
    PC_1_1  = "pc_1_1";  PC_4_2  = "pc_4_2"
    # KBR
    SCORE   = "score"

class ScenarioStatus(str, enum.Enum):
    DRAFT = "draft"; RUNNING = "running"; COMPLETED = "completed"; FAILED = "failed"

class OptimizationStatus(str, enum.Enum):
    QUEUED = "queued"; RUNNING = "running"; OPTIMAL = "optimal"
    INFEASIBLE = "infeasible"; FAILED = "failed"


class TimestampMixin:
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), server_default=func.now(), nullable=False)
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), server_default=func.now(), onupdate=func.now(), nullable=False)

class UUIDPKMixin:
    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)


class Plant(UUIDPKMixin, TimestampMixin, Base):
    __tablename__ = "plants"
    name:        Mapped[str]           = mapped_column(String(100), unique=True, nullable=False)
    code:        Mapped[str]           = mapped_column(String(20),  unique=True, nullable=False)
    location:    Mapped[str]           = mapped_column(String(200), nullable=False)
    capacity_kt: Mapped[float]         = mapped_column(Float, nullable=False)
    status:      Mapped[PlantStatus]   = mapped_column(SAEnum(PlantStatus), default=PlantStatus.ACTIVE)
    description: Mapped[Optional[str]] = mapped_column(Text)
    timezone:    Mapped[str]           = mapped_column(String(50), default="UTC")
    furnaces:    Mapped[List["Furnace"]]  = relationship(back_populates="plant", cascade="all, delete-orphan")
    scenarios:   Mapped[List["Scenario"]] = relationship(back_populates="plant", cascade="all, delete-orphan")
    __table_args__ = (Index("ix_plants_status", "status"),)


class Furnace(UUIDPKMixin, TimestampMixin, Base):
    __tablename__ = "furnaces"
    plant_id:            Mapped[uuid.UUID]     = mapped_column(UUID(as_uuid=True), ForeignKey("plants.id", ondelete="CASCADE"))
    name:                Mapped[str]           = mapped_column(String(100), nullable=False)
    code:                Mapped[str]           = mapped_column(String(20),  nullable=False)
    coil_type:           Mapped[CoilType]      = mapped_column(SAEnum(CoilType), nullable=False)
    status:              Mapped[FurnaceStatus] = mapped_column(SAEnum(FurnaceStatus), default=FurnaceStatus.RUNNING)
    design_feed_rate_th: Mapped[float]         = mapped_column(Float, nullable=False)
    min_feed_rate_th:    Mapped[float]         = mapped_column(Float, nullable=False)
    max_feed_rate_th:    Mapped[float]         = mapped_column(Float, nullable=False)
    efficiency_pct:      Mapped[float]         = mapped_column(Float, default=100.0)
    num_coils:            Mapped[Optional[int]]   = mapped_column(Integer)
    num_passes:           Mapped[Optional[int]]   = mapped_column(Integer)
    feed_rate_source:     Mapped[str]             = mapped_column(String(20), default="manual")
    manual_feed_rate_th:  Mapped[Optional[float]] = mapped_column(Float)
    coil_flow_th:         Mapped[Optional[float]] = mapped_column(Float)
    last_decoking_date:   Mapped[Optional[datetime]] = mapped_column(DateTime(timezone=True))
    notes:               Mapped[Optional[str]] = mapped_column(Text)
    plant:          Mapped["Plant"]             = relationship(back_populates="furnaces")
    operating_data: Mapped[List["FurnaceData"]] = relationship(back_populates="furnace", cascade="all, delete-orphan")
    yield_records:  Mapped[List["YieldRecord"]] = relationship(back_populates="furnace")
    __table_args__ = (
        UniqueConstraint("plant_id", "code", name="uq_furnace_plant_code"),
        Index("ix_furnaces_plant_status", "plant_id", "status"),
    )


class FeedStock(UUIDPKMixin, TimestampMixin, Base):
    __tablename__ = "feedstocks"
    name:                 Mapped[str]            = mapped_column(String(100), unique=True, nullable=False)
    code:                 Mapped[str]            = mapped_column(String(20),  unique=True, nullable=False)
    category:             Mapped[str]            = mapped_column(String(50),  nullable=False)
    density_kg_m3:        Mapped[Optional[float]]= mapped_column(Float)
    molecular_weight:     Mapped[Optional[float]]= mapped_column(Float)
    hydrogen_content_pct: Mapped[Optional[float]]= mapped_column(Float)
    base_price_usd_t:     Mapped[float]          = mapped_column(Float, nullable=False)
    is_active:            Mapped[bool]           = mapped_column(Boolean, default=True)
    yield_records: Mapped[List["YieldRecord"]] = relationship(back_populates="feedstock")
    price_history: Mapped[List["FeedPrice"]]   = relationship(back_populates="feedstock", cascade="all, delete-orphan")


class FeedPrice(UUIDPKMixin, Base):
    __tablename__ = "feed_prices"
    feedstock_id:   Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), ForeignKey("feedstocks.id", ondelete="CASCADE"))
    effective_date: Mapped[datetime]  = mapped_column(DateTime(timezone=True), nullable=False)
    price_usd_t:    Mapped[float]     = mapped_column(Float, nullable=False)
    source:         Mapped[str]       = mapped_column(String(50), default="manual")
    created_at:     Mapped[datetime]  = mapped_column(DateTime(timezone=True), server_default=func.now())
    feedstock: Mapped["FeedStock"] = relationship(back_populates="price_history")
    __table_args__ = (Index("ix_feed_prices_feedstock_date", "feedstock_id", "effective_date"),)


class ProductPrice(UUIDPKMixin, Base):
    __tablename__ = "product_prices"
    product_name:   Mapped[str]      = mapped_column(String(50),  nullable=False)
    effective_date: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)
    price_usd_t:    Mapped[float]    = mapped_column(Float, nullable=False)
    source:         Mapped[str]      = mapped_column(String(50), default="manual")
    created_at:     Mapped[datetime] = mapped_column(DateTime(timezone=True), server_default=func.now())
    __table_args__ = (
        UniqueConstraint("product_name", "effective_date", name="uq_product_price_date"),
        Index("ix_product_prices_name_date", "product_name", "effective_date"),
    )


class FurnaceData(UUIDPKMixin, Base):
    __tablename__ = "furnace_data"
    furnace_id:           Mapped[uuid.UUID]          = mapped_column(UUID(as_uuid=True), ForeignKey("furnaces.id", ondelete="CASCADE"))
    period_date:          Mapped[datetime]            = mapped_column(DateTime(timezone=True), nullable=False)
    feed_rate_th:         Mapped[float]               = mapped_column(Float, nullable=False)
    feedstock_id:         Mapped[Optional[uuid.UUID]] = mapped_column(UUID(as_uuid=True), ForeignKey("feedstocks.id"))
    cot_celsius:          Mapped[Optional[float]]     = mapped_column(Float)
    dilution_steam_ratio: Mapped[Optional[float]]     = mapped_column(Float)
    runtime_hours:        Mapped[Optional[float]]     = mapped_column(Float)
    on_stream_pct:        Mapped[Optional[float]]     = mapped_column(Float)
    created_at:           Mapped[datetime]            = mapped_column(DateTime(timezone=True), server_default=func.now())
    furnace: Mapped["Furnace"] = relationship(back_populates="operating_data")
    __table_args__ = (Index("ix_furnace_data_furnace_period", "furnace_id", "period_date"),)


class YieldRecord(UUIDPKMixin, Base):
    __tablename__ = "yield_records"
    furnace_id:    Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), ForeignKey("furnaces.id", ondelete="CASCADE"))
    feedstock_id:  Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), ForeignKey("feedstocks.id"))
    period_date:   Mapped[datetime]  = mapped_column(DateTime(timezone=True), nullable=False)
    source:        Mapped[str]       = mapped_column(String(20), default="actual")
    ethylene_pct:  Mapped[Optional[float]] = mapped_column(Float)
    propylene_pct: Mapped[Optional[float]] = mapped_column(Float)
    butadiene_pct: Mapped[Optional[float]] = mapped_column(Float)
    benzene_pct:   Mapped[Optional[float]] = mapped_column(Float)
    toluene_pct:   Mapped[Optional[float]] = mapped_column(Float)
    hydrogen_pct:  Mapped[Optional[float]] = mapped_column(Float)
    methane_pct:   Mapped[Optional[float]] = mapped_column(Float)
    fuel_oil_pct:  Mapped[Optional[float]] = mapped_column(Float)
    pygas_pct:     Mapped[Optional[float]] = mapped_column(Float)
    c4_pct:        Mapped[Optional[float]] = mapped_column(Float)
    cot_celsius:   Mapped[Optional[float]] = mapped_column(Float)
    created_at:    Mapped[datetime]        = mapped_column(DateTime(timezone=True), server_default=func.now())
    furnace:   Mapped["Furnace"]   = relationship(back_populates="yield_records")
    feedstock: Mapped["FeedStock"] = relationship(back_populates="yield_records")
    __table_args__ = (Index("ix_yield_furnace_feed_date", "furnace_id", "feedstock_id", "period_date"),)


class Scenario(UUIDPKMixin, TimestampMixin, Base):
    __tablename__ = "scenarios"
    plant_id:            Mapped[uuid.UUID]       = mapped_column(UUID(as_uuid=True), ForeignKey("plants.id", ondelete="CASCADE"))
    name:                Mapped[str]             = mapped_column(String(200), nullable=False)
    description:         Mapped[Optional[str]]   = mapped_column(Text)
    period_start:        Mapped[datetime]         = mapped_column(DateTime(timezone=True), nullable=False)
    period_end:          Mapped[datetime]         = mapped_column(DateTime(timezone=True), nullable=False)
    status:              Mapped[ScenarioStatus]   = mapped_column(SAEnum(ScenarioStatus), default=ScenarioStatus.DRAFT)
    feed_constraints:    Mapped[Optional[dict]]   = mapped_column(JSON)
    furnace_constraints: Mapped[Optional[dict]]   = mapped_column(JSON)
    product_targets:     Mapped[Optional[dict]]   = mapped_column(JSON)
    price_assumptions:   Mapped[Optional[dict]]   = mapped_column(JSON)
    optimization_result: Mapped[Optional[dict]]   = mapped_column(JSON)
    total_margin_usd:    Mapped[Optional[float]]  = mapped_column(Float)
    solver_used:         Mapped[Optional[str]]    = mapped_column(String(50))
    created_by:          Mapped[Optional[str]]    = mapped_column(String(100))
    plant:         Mapped["Plant"]                 = relationship(back_populates="scenarios")
    optimizations: Mapped[List["OptimizationRun"]] = relationship(back_populates="scenario", cascade="all, delete-orphan")
    __table_args__ = (Index("ix_scenarios_plant_status", "plant_id", "status"),)


class OptimizationRun(UUIDPKMixin, Base):
    __tablename__ = "optimization_runs"
    scenario_id:     Mapped[uuid.UUID]          = mapped_column(UUID(as_uuid=True), ForeignKey("scenarios.id", ondelete="CASCADE"))
    celery_task_id:  Mapped[Optional[str]]      = mapped_column(String(200))
    solver_type:     Mapped[Optional[str]]      = mapped_column(String(50), default="minlp_pyomo")
    status:          Mapped[OptimizationStatus] = mapped_column(SAEnum(OptimizationStatus), default=OptimizationStatus.QUEUED)
    started_at:      Mapped[Optional[datetime]] = mapped_column(DateTime(timezone=True))
    completed_at:    Mapped[Optional[datetime]] = mapped_column(DateTime(timezone=True))
    solve_time_sec:  Mapped[Optional[float]]    = mapped_column(Float)
    solver_status:   Mapped[Optional[str]]      = mapped_column(String(200))
    objective_value: Mapped[Optional[float]]    = mapped_column(Float)
    result_detail:   Mapped[Optional[dict]]     = mapped_column(JSON)
    error_message:   Mapped[Optional[str]]      = mapped_column(Text)
    created_at:      Mapped[datetime]           = mapped_column(DateTime(timezone=True), server_default=func.now())
    scenario: Mapped["Scenario"] = relationship(back_populates="optimizations")


class UploadRecord(UUIDPKMixin, Base):
    __tablename__ = "upload_records"
    filename:        Mapped[str]            = mapped_column(String(500), nullable=False)
    original_name:   Mapped[str]            = mapped_column(String(500), nullable=False)
    file_size_bytes: Mapped[int]            = mapped_column(Integer, nullable=False)
    upload_type:     Mapped[str]            = mapped_column(String(50), nullable=False)
    plant_id:        Mapped[Optional[uuid.UUID]] = mapped_column(UUID(as_uuid=True), ForeignKey("plants.id"))
    status:          Mapped[str]            = mapped_column(String(20), default="processing")
    rows_processed:  Mapped[Optional[int]]  = mapped_column(Integer)
    rows_failed:     Mapped[Optional[int]]  = mapped_column(Integer)
    error_detail:    Mapped[Optional[dict]] = mapped_column(JSON)
    created_at:      Mapped[datetime]       = mapped_column(DateTime(timezone=True), server_default=func.now())
    completed_at:    Mapped[Optional[datetime]] = mapped_column(DateTime(timezone=True))


class SensitivityData(UUIDPKMixin, Base):
    __tablename__ = "sensitivity_data"
    furnace_id:   Mapped[str]            = mapped_column(String(50), nullable=False)
    plant_code:   Mapped[Optional[str]]  = mapped_column(String(20))
    feed_type:    Mapped[Optional[str]]  = mapped_column(String(50))
    coil_type:    Mapped[Optional[str]]  = mapped_column(String(50))
    cip:          Mapped[Optional[float]]= mapped_column(Float)
    severity:     Mapped[Optional[float]]= mapped_column(Float)
    feed_rate:    Mapped[Optional[float]]= mapped_column(Float)
    shc:          Mapped[Optional[float]]= mapped_column(Float)
    cit:          Mapped[Optional[float]]= mapped_column(Float)
    cop:          Mapped[Optional[float]]= mapped_column(Float)
    feed_ethane:  Mapped[Optional[float]]= mapped_column(Float)
    feed_propane: Mapped[Optional[float]]= mapped_column(Float)
    cot:          Mapped[Optional[float]]= mapped_column(Float)
    hydrogen_yield:  Mapped[Optional[float]]= mapped_column(Float)
    methane_yield:   Mapped[Optional[float]]= mapped_column(Float)
    coking_rate:     Mapped[Optional[float]]= mapped_column(Float)
    heat_absorbed:   Mapped[Optional[float]]= mapped_column(Float)
    ethylene_yield:  Mapped[Optional[float]]= mapped_column(Float)
    c2h6_yield:      Mapped[Optional[float]]= mapped_column(Float)
    benzene_exit:    Mapped[Optional[float]]= mapped_column(Float)
    styrene_exit:    Mapped[Optional[float]]= mapped_column(Float)
    c2h2_exit:       Mapped[Optional[float]]= mapped_column(Float)
    propylene_yield: Mapped[Optional[float]]= mapped_column(Float)
    dilution_steam:  Mapped[Optional[float]]= mapped_column(Float)
    upload_batch_id: Mapped[Optional[uuid.UUID]] = mapped_column(UUID(as_uuid=True))
    uploaded_at:     Mapped[datetime]       = mapped_column(DateTime(timezone=True), server_default=func.now())
    __table_args__ = (Index("ix_sensitivity_data_furnace", "furnace_id"),)


class SensitivityModel(UUIDPKMixin, Base):
    __tablename__ = "sensitivity_models"
    furnace_id:      Mapped[Optional[str]]  = mapped_column(String(50))
    feed_type:       Mapped[Optional[str]]  = mapped_column(String(50))
    coil_type:       Mapped[Optional[str]]  = mapped_column(String(50))
    model_type:      Mapped[str]            = mapped_column(String(20),  nullable=False)
    degree:          Mapped[int]            = mapped_column(Integer, default=1)
    output_variable: Mapped[str]            = mapped_column(String(50),  nullable=False)
    coefficients:    Mapped[Optional[dict]] = mapped_column(JSON)
    intercept:       Mapped[Optional[float]]= mapped_column(Float)
    feature_names:   Mapped[Optional[dict]] = mapped_column(JSON)
    scaler_mean:     Mapped[Optional[dict]] = mapped_column(JSON)
    scaler_std:      Mapped[Optional[dict]] = mapped_column(JSON)
    r2:              Mapped[Optional[float]]= mapped_column(Float)
    adj_r2:          Mapped[Optional[float]]= mapped_column(Float)
    rmse:            Mapped[Optional[float]]= mapped_column(Float)
    cv_score:        Mapped[Optional[float]]= mapped_column(Float)
    is_best_model:   Mapped[bool]           = mapped_column(Boolean, default=False)
    is_active:       Mapped[bool]           = mapped_column(Boolean, default=False)
    n_samples:       Mapped[Optional[int]]  = mapped_column(Integer)
    actuals:         Mapped[Optional[dict]]  = mapped_column(JSON)
    predictions:     Mapped[Optional[dict]]  = mapped_column(JSON)
    model_artifact:  Mapped[Optional[bytes]] = mapped_column(LargeBinary)
    created_at:      Mapped[datetime]        = mapped_column(DateTime(timezone=True), server_default=func.now())
    __table_args__ = (Index("ix_sensitivity_models_furnace_output", "furnace_id", "output_variable"),)


class RegressionRun(UUIDPKMixin, Base):
    __tablename__ = "regression_runs"
    plant_id:      Mapped[uuid.UUID]          = mapped_column(UUID(as_uuid=True), ForeignKey("plants.id", ondelete="CASCADE"))
    status:        Mapped[str]                = mapped_column(String(20), default="pending")
    n_samples:     Mapped[Optional[int]]      = mapped_column(Integer)
    n_models:      Mapped[Optional[int]]      = mapped_column(Integer)
    error_message: Mapped[Optional[str]]      = mapped_column(Text)
    started_at:    Mapped[Optional[datetime]] = mapped_column(DateTime(timezone=True))
    completed_at:  Mapped[Optional[datetime]] = mapped_column(DateTime(timezone=True))
    created_at:    Mapped[datetime]           = mapped_column(DateTime(timezone=True), server_default=func.now())
    models: Mapped[List["RegressionModel"]] = relationship(back_populates="run", cascade="all, delete-orphan")
    __table_args__ = (Index("ix_regression_runs_plant", "plant_id", "status"),)


class RegressionModel(UUIDPKMixin, Base):
    __tablename__ = "regression_models"
    run_id:          Mapped[uuid.UUID]       = mapped_column(UUID(as_uuid=True), ForeignKey("regression_runs.id", ondelete="CASCADE"))
    plant_id:        Mapped[uuid.UUID]       = mapped_column(UUID(as_uuid=True), ForeignKey("plants.id", ondelete="CASCADE"))
    target_variable: Mapped[str]             = mapped_column(String(50),  nullable=False)
    model_type:      Mapped[str]             = mapped_column(String(50),  nullable=False)
    feature_names:   Mapped[Optional[dict]]  = mapped_column(JSON)
    coefficients:    Mapped[Optional[dict]]  = mapped_column(JSON)
    r2_score:        Mapped[Optional[float]] = mapped_column(Float)
    adj_r2:          Mapped[Optional[float]] = mapped_column(Float)
    rmse:            Mapped[Optional[float]] = mapped_column(Float)
    mae:             Mapped[Optional[float]] = mapped_column(Float)
    mape:            Mapped[Optional[float]] = mapped_column(Float)
    cv_r2:           Mapped[Optional[float]] = mapped_column(Float)
    is_best:         Mapped[bool]            = mapped_column(Boolean, default=False)
    n_samples:       Mapped[Optional[int]]   = mapped_column(Integer)
    created_at:      Mapped[datetime]        = mapped_column(DateTime(timezone=True), server_default=func.now())
    run: Mapped["RegressionRun"] = relationship(back_populates="models")
    __table_args__ = (Index("ix_regression_models_plant_target", "plant_id", "target_variable"),)`;

const SCHEMAS_PY = `import uuid
from datetime import datetime
from typing import Optional, List, Dict, Any
from pydantic import BaseModel, Field, ConfigDict
from app.models.models import PlantStatus, FurnaceStatus, CoilType, ScenarioStatus, OptimizationStatus


class OrmBase(BaseModel):
    model_config = ConfigDict(from_attributes=True)


# -- Plant -------------------------------------------------------------------
class PlantCreate(BaseModel):
    name: str = Field(..., min_length=2, max_length=100)
    code: str = Field(..., min_length=2, max_length=20)
    location: str
    capacity_kt: float = Field(..., gt=0)
    status: PlantStatus = PlantStatus.ACTIVE
    description: Optional[str] = None
    timezone: str = "UTC"

class PlantUpdate(BaseModel):
    name: Optional[str] = None
    location: Optional[str] = None
    capacity_kt: Optional[float] = Field(None, gt=0)
    status: Optional[PlantStatus] = None
    description: Optional[str] = None

class PlantResponse(OrmBase):
    id: uuid.UUID; name: str; code: str; location: str
    capacity_kt: float; status: PlantStatus; description: Optional[str]
    timezone: str; created_at: datetime; updated_at: datetime
    furnace_count: Optional[int] = None


# -- Furnace -----------------------------------------------------------------
class FurnaceCreate(BaseModel):
    plant_id: uuid.UUID
    name: str = Field(..., min_length=2)
    code: str = Field(..., min_length=1, max_length=20)
    coil_type: CoilType
    design_feed_rate_th: float = Field(..., gt=0)
    min_feed_rate_th: float = Field(..., gt=0)
    max_feed_rate_th: float = Field(..., gt=0)
    efficiency_pct: float = Field(100.0, ge=0, le=110)
    num_coils:  Optional[int] = Field(None, ge=1, le=100)
    num_passes: Optional[int] = Field(None, ge=1, le=32)
    feed_rate_source:    str            = Field("manual", pattern="^(manual|coil_calc|data_link)$")
    manual_feed_rate_th: Optional[float]= Field(None, gt=0)
    coil_flow_th:        Optional[float]= Field(None, gt=0)
    notes: Optional[str] = None

class FurnaceUpdate(BaseModel):
    coil_type: Optional[CoilType] = None
    status: Optional[FurnaceStatus] = None
    design_feed_rate_th: Optional[float] = Field(None, gt=0)
    min_feed_rate_th: Optional[float] = Field(None, gt=0)
    max_feed_rate_th: Optional[float] = Field(None, gt=0)
    efficiency_pct: Optional[float] = Field(None, ge=0, le=110)
    num_coils:  Optional[int] = Field(None, ge=1, le=100)
    num_passes: Optional[int] = Field(None, ge=1, le=32)
    feed_rate_source:    Optional[str]   = Field(None, pattern="^(manual|coil_calc|data_link)$")
    manual_feed_rate_th: Optional[float] = Field(None, gt=0)
    coil_flow_th:        Optional[float] = Field(None, gt=0)
    last_decoking_date: Optional[datetime] = None
    notes: Optional[str] = None

class FurnaceResponse(OrmBase):
    id: uuid.UUID; plant_id: uuid.UUID; name: str; code: str
    coil_type: CoilType; status: FurnaceStatus
    design_feed_rate_th: float; min_feed_rate_th: float; max_feed_rate_th: float
    efficiency_pct: float; num_coils: Optional[int]; num_passes: Optional[int]
    feed_rate_source: str
    manual_feed_rate_th: Optional[float]; coil_flow_th: Optional[float]
    actual_feed_rate_th: Optional[float] = None   # computed by route, not a DB column
    last_decoking_date: Optional[datetime]; notes: Optional[str]
    created_at: datetime


# -- FeedStock ---------------------------------------------------------------
class FeedStockCreate(BaseModel):
    name: str; code: str; category: str
    base_price_usd_t: float = Field(..., gt=0)
    density_kg_m3: Optional[float] = None
    molecular_weight: Optional[float] = None
    hydrogen_content_pct: Optional[float] = Field(None, ge=0, le=100)

class FeedStockResponse(OrmBase):
    id: uuid.UUID; name: str; code: str; category: str
    base_price_usd_t: float; density_kg_m3: Optional[float]
    is_active: bool; created_at: datetime


# -- Yield -------------------------------------------------------------------
class YieldRecordCreate(BaseModel):
    furnace_id: uuid.UUID; feedstock_id: uuid.UUID; period_date: datetime
    source: str = "actual"
    ethylene_pct: Optional[float] = Field(None, ge=0, le=100)
    propylene_pct: Optional[float] = Field(None, ge=0, le=100)
    butadiene_pct: Optional[float] = Field(None, ge=0, le=100)
    benzene_pct: Optional[float] = Field(None, ge=0, le=100)
    toluene_pct: Optional[float] = Field(None, ge=0, le=100)
    hydrogen_pct: Optional[float] = Field(None, ge=0, le=100)
    methane_pct: Optional[float] = Field(None, ge=0, le=100)
    fuel_oil_pct: Optional[float] = Field(None, ge=0, le=100)
    pygas_pct: Optional[float] = Field(None, ge=0, le=100)
    c4_pct: Optional[float] = Field(None, ge=0, le=100)
    cot_celsius: Optional[float] = None


# -- Scenario ----------------------------------------------------------------
class ScenarioCreate(BaseModel):
    plant_id: uuid.UUID
    name: str = Field(..., min_length=2)
    description: Optional[str] = None
    period_start: datetime; period_end: datetime
    feed_constraints: Optional[Dict[str, Any]] = None
    furnace_constraints: Optional[Dict[str, Any]] = None
    product_targets: Optional[Dict[str, Any]] = None
    price_assumptions: Optional[Dict[str, Any]] = None
    created_by: Optional[str] = None

class ScenarioUpdate(BaseModel):
    name: Optional[str] = None; description: Optional[str] = None
    status: Optional[ScenarioStatus] = None
    feed_constraints: Optional[Dict[str, Any]] = None
    furnace_constraints: Optional[Dict[str, Any]] = None
    product_targets: Optional[Dict[str, Any]] = None
    price_assumptions: Optional[Dict[str, Any]] = None

class ScenarioResponse(OrmBase):
    id: uuid.UUID; plant_id: uuid.UUID; name: str; description: Optional[str]
    period_start: datetime; period_end: datetime; status: ScenarioStatus
    feed_constraints: Optional[Dict[str, Any]]; furnace_constraints: Optional[Dict[str, Any]]
    product_targets: Optional[Dict[str, Any]]; price_assumptions: Optional[Dict[str, Any]]
    optimization_result: Optional[Dict[str, Any]]; total_margin_usd: Optional[float]
    solver_used: Optional[str]; created_by: Optional[str]
    created_at: datetime; updated_at: datetime


# -- Optimization ------------------------------------------------------------
class OptimizationRequest(BaseModel):
    scenario_id: uuid.UUID
    solver: str = "minlp_pyomo"

class OptRunResponse(OrmBase):
    id: uuid.UUID; scenario_id: uuid.UUID; celery_task_id: Optional[str]
    solver_type: Optional[str]; status: OptimizationStatus
    started_at: Optional[datetime]; completed_at: Optional[datetime]
    solve_time_sec: Optional[float]; solver_status: Optional[str]
    objective_value: Optional[float]; result_detail: Optional[Dict[str, Any]]
    error_message: Optional[str]; created_at: datetime`;

const HEALTH_PY = `"""health.py"""
from datetime import datetime, timezone
from fastapi import APIRouter
router = APIRouter()

@router.get("/health")
async def health():
    return {"status": "ok", "timestamp": datetime.now(timezone.utc).isoformat()}`;

const DASHBOARD_PY = `from datetime import datetime, timezone, timedelta
from fastapi import APIRouter, Depends
from sqlalchemy import select, func
from sqlalchemy.ext.asyncio import AsyncSession
from app.core.database import get_db
from app.models.models import Plant, Furnace, FurnaceData, Scenario, ProductPrice, FeedPrice, PlantStatus, FurnaceStatus

router = APIRouter()


@router.get("/summary")
async def summary(db: AsyncSession = Depends(get_db)):
    plants   = (await db.execute(select(Plant))).scalars().all()
    furnaces = (await db.execute(select(Furnace))).scalars().all()
    running  = [f for f in furnaces if f.status == FurnaceStatus.RUNNING]

    since  = datetime.now(timezone.utc) - timedelta(hours=24)
    fd_res = await db.execute(select(FurnaceData).where(FurnaceData.period_date >= since))
    recent = fd_res.scalars().all()
    feed_rate = (sum(r.feed_rate_th for r in recent) / len(recent)) if recent else \\
                sum(f.design_feed_rate_th for f in running)

    sc_res = await db.execute(
        select(func.count(Scenario.id)).where(Scenario.status.in_(["draft","running"]))
    )
    active_scenarios = sc_res.scalar() or 0

    margin_res = await db.execute(
        select(Scenario).where(Scenario.total_margin_usd != None,
                                Scenario.status == "completed")
        .order_by(Scenario.updated_at.desc()).limit(10)
    )
    active_plants = [p for p in plants if p.status == PlantStatus.ACTIVE]
    total_margin  = sum(s.total_margin_usd or 0 for s in margin_res.scalars())

    return {
        "total_plants": len(plants), "active_plants": len(active_plants),
        "total_furnaces": len(furnaces), "running_furnaces": len(running),
        "total_feed_rate_th": round(feed_rate, 2),
        "active_scenarios": active_scenarios,
        "total_margin_usd_day": round(total_margin, 0),
        "last_updated": datetime.now(timezone.utc).isoformat(),
    }


@router.get("/plants-kpi")
async def plants_kpi(db: AsyncSession = Depends(get_db)):
    plants = (await db.execute(select(Plant).order_by(Plant.name))).scalars().all()
    out = []
    for plant in plants:
        furnaces = (await db.execute(select(Furnace).where(Furnace.plant_id == plant.id))).scalars().all()
        running  = [f for f in furnaces if f.status == FurnaceStatus.RUNNING]
        max_rate = sum(f.max_feed_rate_th for f in running)
        design   = sum(f.design_feed_rate_th for f in running)

        sc = (await db.execute(
            select(Scenario).where(Scenario.plant_id == plant.id, Scenario.total_margin_usd != None)
            .order_by(Scenario.updated_at.desc()).limit(1)
        )).scalar_one_or_none()

        out.append({
            "plant_id": str(plant.id), "plant_name": plant.name, "plant_code": plant.code,
            "location": plant.location, "status": plant.status.value,
            "furnaces_running": len(running), "furnaces_total": len(furnaces),
            "feed_rate_th": round(max_rate, 2),
            "capacity_util_pct": round(max_rate / max(design, 1) * 100, 1) if design else 0,
            "ethylene_rate_th": round(max_rate * 0.30, 2),
            "daily_margin_usd": round(sc.total_margin_usd or 0, 0) if sc else 0,
        })
    return out


@router.get("/furnace-status")
async def furnace_status(db: AsyncSession = Depends(get_db)):
    r = await db.execute(
        select(Furnace, Plant.name.label("plant_name"), Plant.code.label("plant_code"))
        .join(Plant, Plant.id == Furnace.plant_id)
        .order_by(Plant.name, Furnace.code)
    )
    return [
        {
            "furnace_id": str(row.Furnace.id), "furnace_name": row.Furnace.name,
            "furnace_code": row.Furnace.code, "plant_name": row.plant_name,
            "plant_code": row.plant_code, "coil_type": row.Furnace.coil_type.value,
            "status": row.Furnace.status.value, "max_rate_th": row.Furnace.max_feed_rate_th,
            "efficiency_pct": row.Furnace.efficiency_pct,
        }
        for row in r.all()
    ]


@router.get("/product-prices")
async def product_prices(db: AsyncSession = Depends(get_db)):
    r = await db.execute(select(ProductPrice).order_by(
        ProductPrice.product_name, ProductPrice.effective_date.desc()))
    seen = set(); out = []
    for p in r.scalars():
        if p.product_name not in seen:
            out.append({"product": p.product_name, "price_usd_t": p.price_usd_t,
                         "date": p.effective_date.isoformat(), "source": p.source})
            seen.add(p.product_name)
    return out`;

const PLANTS_PY = `import uuid
from typing import List, Optional
from fastapi import APIRouter, Depends, HTTPException, Query
from sqlalchemy import select, func
from sqlalchemy.ext.asyncio import AsyncSession
from app.core.database import get_db
from app.models.models import Plant, Furnace
from app.schemas.schemas import PlantCreate, PlantUpdate, PlantResponse

router = APIRouter()

@router.get("/", response_model=List[PlantResponse])
async def list_plants(status: Optional[str] = Query(None), db: AsyncSession = Depends(get_db)):
    stmt = select(Plant)
    if status: stmt = stmt.where(Plant.status == status)
    result = await db.execute(stmt.order_by(Plant.name))
    plants = result.scalars().all()
    out = []
    for p in plants:
        cnt = (await db.execute(select(func.count(Furnace.id)).where(Furnace.plant_id == p.id))).scalar()
        r = PlantResponse.model_validate(p); r.furnace_count = cnt; out.append(r)
    return out

@router.post("/", response_model=PlantResponse, status_code=201)
async def create_plant(body: PlantCreate, db: AsyncSession = Depends(get_db)):
    ex = await db.execute(select(Plant).where(Plant.code == body.code))
    if ex.scalar_one_or_none(): raise HTTPException(400, f"Plant code '{body.code}' already exists")
    p = Plant(**body.model_dump()); db.add(p); await db.flush(); await db.refresh(p); return p

@router.get("/{plant_id}", response_model=PlantResponse)
async def get_plant(plant_id: uuid.UUID, db: AsyncSession = Depends(get_db)):
    p = await db.get(Plant, plant_id)
    if not p: raise HTTPException(404, "Plant not found")
    return p

@router.patch("/{plant_id}", response_model=PlantResponse)
async def update_plant(plant_id: uuid.UUID, body: PlantUpdate, db: AsyncSession = Depends(get_db)):
    p = await db.get(Plant, plant_id)
    if not p: raise HTTPException(404, "Plant not found")
    for k, v in body.model_dump(exclude_none=True).items(): setattr(p, k, v)
    await db.flush(); await db.refresh(p); return p

@router.delete("/{plant_id}", status_code=204)
async def delete_plant(plant_id: uuid.UUID, db: AsyncSession = Depends(get_db)):
    p = await db.get(Plant, plant_id)
    if not p: raise HTTPException(404, "Plant not found")
    await db.delete(p)`;

const FURNACES_PY = `import uuid
from typing import List, Optional
from fastapi import APIRouter, Depends, HTTPException, Query
from sqlalchemy import select, func, and_
from sqlalchemy.ext.asyncio import AsyncSession
from app.core.database import get_db
from app.models.models import Furnace, Plant, FurnaceData
from app.schemas.schemas import FurnaceCreate, FurnaceUpdate, FurnaceResponse

router = APIRouter()


def _compute_actual(f: Furnace, latest_data_rates: dict) -> Optional[float]:
    """Compute actual_feed_rate_th based on feed_rate_source."""
    src = f.feed_rate_source or "manual"
    if src == "manual":
        return f.manual_feed_rate_th
    if src == "coil_calc":
        if f.coil_flow_th and f.num_coils:
            return round(f.coil_flow_th * f.num_coils, 3)
        return None
    if src == "data_link":
        return latest_data_rates.get(str(f.id))
    return None


async def _latest_data_rates(furnace_ids: list[uuid.UUID], db: AsyncSession) -> dict:
    """Return {furnace_id_str: latest feed_rate_th} from furnace_data."""
    if not furnace_ids:
        return {}
    latest_subq = (
        select(FurnaceData.furnace_id, func.max(FurnaceData.period_date).label("max_date"))
        .where(FurnaceData.furnace_id.in_(furnace_ids))
        .group_by(FurnaceData.furnace_id)
        .subquery()
    )
    stmt = (
        select(FurnaceData.furnace_id, FurnaceData.feed_rate_th)
        .join(latest_subq, and_(
            FurnaceData.furnace_id == latest_subq.c.furnace_id,
            FurnaceData.period_date == latest_subq.c.max_date,
        ))
    )
    rows = (await db.execute(stmt)).all()
    return {str(r.furnace_id): r.feed_rate_th for r in rows}


def _to_response(f: Furnace, actual: Optional[float]) -> FurnaceResponse:
    resp = FurnaceResponse.model_validate(f)
    resp.actual_feed_rate_th = actual
    return resp


@router.get("/", response_model=List[FurnaceResponse])
async def list_furnaces(
    plant_id: Optional[uuid.UUID] = Query(None),
    status:   Optional[str]       = Query(None),
    db: AsyncSession = Depends(get_db),
):
    stmt = select(Furnace)
    if plant_id: stmt = stmt.where(Furnace.plant_id == plant_id)
    if status:   stmt = stmt.where(Furnace.status == status)
    furnaces = (await db.execute(stmt.order_by(Furnace.name))).scalars().all()

    # Fetch latest operating data rates for data_link furnaces
    data_link_ids = [f.id for f in furnaces if (f.feed_rate_source or "manual") == "data_link"]
    latest_rates  = await _latest_data_rates(data_link_ids, db)

    return [_to_response(f, _compute_actual(f, latest_rates)) for f in furnaces]


@router.post("/", response_model=FurnaceResponse, status_code=201)
async def create_furnace(body: FurnaceCreate, db: AsyncSession = Depends(get_db)):
    if not await db.get(Plant, body.plant_id):
        raise HTTPException(404, "Plant not found")
    f = Furnace(**body.model_dump())
    db.add(f); await db.flush(); await db.refresh(f)
    actual = _compute_actual(f, {})
    return _to_response(f, actual)


@router.get("/{furnace_id}", response_model=FurnaceResponse)
async def get_furnace(furnace_id: uuid.UUID, db: AsyncSession = Depends(get_db)):
    f = await db.get(Furnace, furnace_id)
    if not f: raise HTTPException(404, "Furnace not found")
    rates  = await _latest_data_rates([f.id], db)
    return _to_response(f, _compute_actual(f, rates))


@router.patch("/{furnace_id}", response_model=FurnaceResponse)
async def update_furnace(furnace_id: uuid.UUID, body: FurnaceUpdate, db: AsyncSession = Depends(get_db)):
    f = await db.get(Furnace, furnace_id)
    if not f: raise HTTPException(404, "Furnace not found")
    for k, v in body.model_dump(exclude_none=True).items():
        setattr(f, k, v)
    await db.flush(); await db.refresh(f)
    rates  = await _latest_data_rates([f.id], db)
    return _to_response(f, _compute_actual(f, rates))


@router.delete("/{furnace_id}", status_code=204)
async def delete_furnace(furnace_id: uuid.UUID, db: AsyncSession = Depends(get_db)):
    f = await db.get(Furnace, furnace_id)
    if not f: raise HTTPException(404, "Furnace not found")
    await db.delete(f)`;

const FEEDS_PY = `import uuid
from typing import List
from fastapi import APIRouter, Depends, HTTPException, Query
from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession
from app.core.database import get_db
from app.models.models import FeedStock, FeedPrice
from app.schemas.schemas import FeedStockCreate, FeedStockResponse

router = APIRouter()

@router.get("/", response_model=List[FeedStockResponse])
async def list_feeds(active_only: bool = Query(True), db: AsyncSession = Depends(get_db)):
    stmt = select(FeedStock)
    if active_only: stmt = stmt.where(FeedStock.is_active == True)
    r = await db.execute(stmt.order_by(FeedStock.name)); return r.scalars().all()

@router.post("/", response_model=FeedStockResponse, status_code=201)
async def create_feed(body: FeedStockCreate, db: AsyncSession = Depends(get_db)):
    ex = await db.execute(select(FeedStock).where(FeedStock.code == body.code))
    if ex.scalar_one_or_none(): raise HTTPException(400, f"Feed '{body.code}' exists")
    f = FeedStock(**body.model_dump()); db.add(f); await db.flush(); await db.refresh(f); return f

@router.get("/{feed_id}", response_model=FeedStockResponse)
async def get_feed(feed_id: uuid.UUID, db: AsyncSession = Depends(get_db)):
    f = await db.get(FeedStock, feed_id)
    if not f: raise HTTPException(404, "Feed not found"); return f

@router.get("/{feed_id}/prices")
async def get_feed_prices(feed_id: uuid.UUID, limit: int = Query(90, le=365), db: AsyncSession = Depends(get_db)):
    r = await db.execute(select(FeedPrice).where(FeedPrice.feedstock_id == feed_id)
                          .order_by(FeedPrice.effective_date.desc()).limit(limit))
    return [{"date": p.effective_date.isoformat(), "price_usd_t": p.price_usd_t, "source": p.source}
            for p in r.scalars()]

@router.post("/{feed_id}/prices", status_code=201)
async def add_feed_price(feed_id: uuid.UUID, price_usd_t: float, effective_date: str,
                         db: AsyncSession = Depends(get_db)):
    from datetime import datetime
    if not await db.get(FeedStock, feed_id): raise HTTPException(404, "Feed not found")
    db.add(FeedPrice(feedstock_id=feed_id, effective_date=datetime.fromisoformat(effective_date),
                     price_usd_t=price_usd_t, source="manual"))
    await db.flush(); return {"ok": True}`;

const YIELD_MATRIX_PY = `import uuid
from datetime import datetime
from typing import List, Optional
from fastapi import APIRouter, Depends, HTTPException, Query
from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession
from app.core.database import get_db
from app.models.models import YieldRecord, Furnace, FeedStock
from app.schemas.schemas import YieldRecordCreate

router = APIRouter()
PRODUCTS = ["ethylene","propylene","butadiene","benzene","toluene","hydrogen","methane","fuel_oil","pygas","c4"]

@router.get("/matrix")
async def get_yield_matrix(plant_id: uuid.UUID = Query(...),
                            date_from: Optional[str] = Query(None), date_to: Optional[str] = Query(None),
                            db: AsyncSession = Depends(get_db)):
    fr = await db.execute(select(Furnace).where(Furnace.plant_id == plant_id))
    furnaces = {str(f.id): f for f in fr.scalars()}
    if not furnaces: raise HTTPException(404, "No furnaces for this plant")

    stmt = select(YieldRecord).where(YieldRecord.furnace_id.in_([uuid.UUID(fid) for fid in furnaces]))
    if date_from: stmt = stmt.where(YieldRecord.period_date >= datetime.fromisoformat(date_from))
    if date_to:   stmt = stmt.where(YieldRecord.period_date <= datetime.fromisoformat(date_to))
    yr = await db.execute(stmt); records = yr.scalars().all()

    kr = await db.execute(select(FeedStock)); feeds = {str(f.id): f for f in kr.scalars()}

    agg: dict = {}
    for r in records:
        fk = str(r.feedstock_id)
        if fk not in agg: agg[fk] = {p: [] for p in PRODUCTS}
        for p in PRODUCTS:
            v = getattr(r, f"{p}_pct", None)
            if v is not None: agg[fk][p].append(v)

    matrix = {
        feeds[fid].name if fid in feeds else fid: {
            p: round(sum(vals)/len(vals), 3) if vals else 0.0 for p, vals in pd.items()
        } for fid, pd in agg.items()
    }
    return {"plant_id": str(plant_id), "record_count": len(records),
            "feedstocks": list(matrix.keys()), "products": PRODUCTS, "matrix": matrix,
            "coil_types": {f.code: f.coil_type.value for f in furnaces.values()}}

@router.get("/records")
async def list_records(furnace_id: Optional[uuid.UUID] = Query(None),
                        feedstock_id: Optional[uuid.UUID] = Query(None),
                        limit: int = Query(200, le=1000), db: AsyncSession = Depends(get_db)):
    stmt = select(YieldRecord)
    if furnace_id:   stmt = stmt.where(YieldRecord.furnace_id == furnace_id)
    if feedstock_id: stmt = stmt.where(YieldRecord.feedstock_id == feedstock_id)
    stmt = stmt.order_by(YieldRecord.period_date.desc()).limit(limit)
    r = await db.execute(stmt)
    return [{"id": str(rec.id), "furnace_id": str(rec.furnace_id), "feedstock_id": str(rec.feedstock_id),
             "period_date": rec.period_date.isoformat(), "source": rec.source,
             **{p: getattr(rec, f"{p}_pct") for p in PRODUCTS},
             "cot_celsius": rec.cot_celsius} for rec in r.scalars()]

@router.post("/records", status_code=201)
async def create_record(body: YieldRecordCreate, db: AsyncSession = Depends(get_db)):
    rec = YieldRecord(**body.model_dump()); db.add(rec); await db.flush()
    return {"id": str(rec.id)}

@router.delete("/records/{record_id}", status_code=204)
async def delete_record(record_id: uuid.UUID, db: AsyncSession = Depends(get_db)):
    r = await db.get(YieldRecord, record_id)
    if not r: raise HTTPException(404, "Record not found")
    await db.delete(r)`;

const SCENARIOS_PY = `import uuid
from typing import List, Optional
from fastapi import APIRouter, Depends, HTTPException, Query
from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession
from app.core.database import get_db
from app.models.models import Scenario, ScenarioStatus, Plant
from app.schemas.schemas import ScenarioCreate, ScenarioUpdate, ScenarioResponse

router = APIRouter()


@router.get("/", response_model=List[ScenarioResponse])
async def list_scenarios(
    plant_id: Optional[uuid.UUID] = Query(None),
    status:   Optional[str]       = Query(None),
    limit:    int                  = Query(50, le=200),
    db: AsyncSession = Depends(get_db),
):
    stmt = select(Scenario)
    if plant_id: stmt = stmt.where(Scenario.plant_id == plant_id)
    if status:   stmt = stmt.where(Scenario.status == status)
    r = await db.execute(stmt.order_by(Scenario.created_at.desc()).limit(limit))
    return r.scalars().all()


@router.post("/", response_model=ScenarioResponse, status_code=201)
async def create_scenario(body: ScenarioCreate, db: AsyncSession = Depends(get_db)):
    if not await db.get(Plant, body.plant_id):
        raise HTTPException(404, "Plant not found")
    if body.period_end <= body.period_start:
        raise HTTPException(400, "period_end must be after period_start")
    s = Scenario(**body.model_dump())
    db.add(s); await db.flush(); await db.refresh(s); return s


@router.get("/compare")
async def compare_scenarios(
    ids: str = Query(..., description="Comma-separated scenario UUIDs"),
    db: AsyncSession = Depends(get_db),
):
    id_list = [uuid.UUID(i.strip()) for i in ids.split(",")]
    r = await db.execute(select(Scenario).where(Scenario.id.in_(id_list)))
    scenarios = r.scalars().all()
    return [
        {
            "scenario_id":    str(s.id),
            "name":           s.name,
            "status":         s.status.value,
            "total_margin_usd": s.total_margin_usd,
            "solver_used":    s.solver_used,
            "period_start":   s.period_start.isoformat(),
            "period_end":     s.period_end.isoformat(),
            "margin_detail":  s.optimization_result.get("margin_detail") if s.optimization_result else None,
        }
        for s in scenarios
    ]


@router.get("/{scenario_id}", response_model=ScenarioResponse)
async def get_scenario(scenario_id: uuid.UUID, db: AsyncSession = Depends(get_db)):
    s = await db.get(Scenario, scenario_id)
    if not s: raise HTTPException(404, "Scenario not found")
    return s


@router.patch("/{scenario_id}", response_model=ScenarioResponse)
async def update_scenario(
    scenario_id: uuid.UUID, body: ScenarioUpdate, db: AsyncSession = Depends(get_db)
):
    s = await db.get(Scenario, scenario_id)
    if not s: raise HTTPException(404, "Scenario not found")
    for k, v in body.model_dump(exclude_none=True).items(): setattr(s, k, v)
    await db.flush(); await db.refresh(s); return s


@router.delete("/{scenario_id}", status_code=204)
async def delete_scenario(scenario_id: uuid.UUID, db: AsyncSession = Depends(get_db)):
    s = await db.get(Scenario, scenario_id)
    if not s: raise HTTPException(404, "Scenario not found")
    await db.delete(s)


@router.post("/{scenario_id}/duplicate", response_model=ScenarioResponse, status_code=201)
async def duplicate_scenario(
    scenario_id: uuid.UUID, new_name: str = Query(...), db: AsyncSession = Depends(get_db)
):
    s = await db.get(Scenario, scenario_id)
    if not s: raise HTTPException(404, "Scenario not found")
    clone = Scenario(
        plant_id=s.plant_id, name=new_name,
        description=f"Copy of: {s.name}",
        period_start=s.period_start, period_end=s.period_end,
        status=ScenarioStatus.DRAFT,
        feed_constraints=s.feed_constraints,
        furnace_constraints=s.furnace_constraints,
        product_targets=s.product_targets,
        price_assumptions=s.price_assumptions,
        created_by=s.created_by,
    )
    db.add(clone); await db.flush(); await db.refresh(clone); return clone`;

const OPTIMIZATION_PY = `import uuid
from typing import List
from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession
from app.core.database import get_db
from app.models.models import OptimizationRun, Scenario, OptimizationStatus
from app.schemas.schemas import OptimizationRequest, OptRunResponse

router = APIRouter()


@router.post("/run", response_model=OptRunResponse, status_code=202)
async def trigger_optimization(body: OptimizationRequest, db: AsyncSession = Depends(get_db)):
    """Kick off a background MINLP solve for a scenario. Returns immediately."""
    if not await db.get(Scenario, body.scenario_id):
        raise HTTPException(404, "Scenario not found")

    run = OptimizationRun(scenario_id=body.scenario_id, status=OptimizationStatus.QUEUED,
                          solver_type=body.solver)
    db.add(run); await db.flush(); await db.refresh(run)

    # Dispatch to Celery; fall back to sync execution if Celery unavailable
    try:
        from app.services.optimization_tasks import run_optimization_task
        task = run_optimization_task.delay(str(run.id))
        run.celery_task_id = task.id
        await db.flush()
    except Exception:
        # Synchronous fallback for environments without Redis
        run.celery_task_id = "sync"
        await db.commit()
        from app.services.optimization_tasks import _run
        import asyncio
        asyncio.create_task(_run(str(run.id), "sync"))

    return run


@router.get("/runs/{run_id}", response_model=OptRunResponse)
async def get_run(run_id: uuid.UUID, db: AsyncSession = Depends(get_db)):
    run = await db.get(OptimizationRun, run_id)
    if not run: raise HTTPException(404, "Optimization run not found")
    return run


@router.get("/scenario/{scenario_id}/runs", response_model=List[OptRunResponse])
async def list_runs(scenario_id: uuid.UUID, db: AsyncSession = Depends(get_db)):
    r = await db.execute(
        select(OptimizationRun)
        .where(OptimizationRun.scenario_id == scenario_id)
        .order_by(OptimizationRun.created_at.desc())
    )
    return r.scalars().all()


@router.get("/scenario/{scenario_id}/latest")
async def latest_result(scenario_id: uuid.UUID, db: AsyncSession = Depends(get_db)):
    r = await db.execute(
        select(OptimizationRun)
        .where(OptimizationRun.scenario_id == scenario_id,
               OptimizationRun.status == OptimizationStatus.OPTIMAL)
        .order_by(OptimizationRun.completed_at.desc())
        .limit(1)
    )
    run = r.scalar_one_or_none()
    if not run: raise HTTPException(404, "No completed optimization for this scenario")
    return {
        "run_id": str(run.id),
        "objective_usd": run.objective_value,
        "solver_used": run.solver_type,
        "solve_time_sec": run.solve_time_sec,
        "completed_at": run.completed_at.isoformat() if run.completed_at else None,
        **(run.result_detail or {}),
    }`;

const UPLOADS_PY = `import os, uuid
from typing import Optional
from fastapi import APIRouter, Depends, File, Form, HTTPException, UploadFile
from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession
from app.core.config import settings
from app.core.database import get_db
from app.models.models import UploadRecord
from app.services.excel_service import ExcelUploadService

router = APIRouter()
ALLOWED_TYPES = {"yield_data", "feed_prices", "product_prices", "operating_data"}
ALLOWED_EXTS  = {".xlsx", ".xlsm", ".xls", ".csv"}


@router.post("/")
async def upload_file(
    file:        UploadFile = File(...),
    upload_type: str        = Form(...),
    plant_id:    Optional[str] = Form(None),
    db: AsyncSession = Depends(get_db),
):
    if upload_type not in ALLOWED_TYPES:
        raise HTTPException(400, f"upload_type must be one of {ALLOWED_TYPES}")

    ext = os.path.splitext(file.filename or "")[1].lower()
    if ext not in ALLOWED_EXTS:
        raise HTTPException(400, f"File must be one of {ALLOWED_EXTS}")

    content = await file.read()
    if len(content) > settings.MAX_UPLOAD_MB * 1024 * 1024:
        raise HTTPException(413, f"File exceeds {settings.MAX_UPLOAD_MB} MB limit")

    os.makedirs(settings.UPLOAD_DIR, exist_ok=True)
    safe_name = f"{uuid.uuid4()}{ext}"
    with open(os.path.join(settings.UPLOAD_DIR, safe_name), "wb") as fh:
        fh.write(content)

    plant_uuid = uuid.UUID(plant_id) if plant_id else None
    record = UploadRecord(
        filename=safe_name, original_name=file.filename or "unknown",
        file_size_bytes=len(content), upload_type=upload_type,
        plant_id=plant_uuid, status="processing",
    )
    db.add(record); await db.flush()

    svc = ExcelUploadService(db)
    try:
        if upload_type == "yield_data":
            ok, fail, errors = await svc.process_yield(content, file.filename, plant_uuid)
        elif upload_type == "feed_prices":
            ok, fail, errors = await svc.process_feed_prices(content, file.filename)
        elif upload_type == "product_prices":
            ok, fail, errors = await svc.process_product_prices(content, file.filename)
        else:
            ok, fail, errors = await svc.process_operating(content, file.filename, plant_uuid)

        from datetime import datetime, timezone
        record.status        = "completed" if fail == 0 else "partial"
        record.rows_processed = ok
        record.rows_failed    = fail
        record.error_detail   = {"errors": errors[:50]} if errors else None
        record.completed_at   = datetime.now(timezone.utc)
        await db.flush()

        return {
            "upload_id": str(record.id), "filename": file.filename,
            "status": record.status, "rows_processed": ok, "rows_failed": fail,
            "errors": errors[:20],
            "message": f"Imported {ok} rows" + (f", {fail} failed" if fail else ""),
        }
    except Exception as e:
        from datetime import datetime, timezone
        record.status = "failed"; record.error_detail = {"error": str(e)}
        record.completed_at = datetime.now(timezone.utc)
        await db.flush()
        raise HTTPException(422, f"Processing failed: {e}")


@router.get("/history")
async def upload_history(limit: int = 50, db: AsyncSession = Depends(get_db)):
    r = await db.execute(
        select(UploadRecord).order_by(UploadRecord.created_at.desc()).limit(limit)
    )
    return [
        {
            "id": str(rec.id), "original_name": rec.original_name,
            "upload_type": rec.upload_type, "status": rec.status,
            "rows_processed": rec.rows_processed, "rows_failed": rec.rows_failed,
            "file_size_kb": round(rec.file_size_bytes / 1024, 1),
            "created_at": rec.created_at.isoformat(),
        }
        for rec in r.scalars()
    ]`;

// Read the remaining files from the filesystem
const BASE = "C:/Users/ptawade/Desktop/SNOP Product/snop/backend/app";
function readFile(relPath) {
  return fs.readFileSync(path.join(BASE, relPath), "utf8");
}

const REGRESSION_PY    = readFile("api/routes/regression.py");
const SENSITIVITY_PY   = readFile("api/routes/sensitivity.py");
const EXCEL_SVC_PY     = readFile("services/excel_service.py");
const OPTIMIZER_PY     = readFile("services/optimizer.py");
const OPT_TASKS_PY     = readFile("services/optimization_tasks.py");
const REG_SVC_PY       = readFile("services/regression_service.py");
const SENS_REG_SVC_PY  = readFile("services/sensitivity_regression_service.py");
const SENS_UPL_SVC_PY  = readFile("services/sensitivity_upload_service.py");

// ── build document children ──────────────────────────────────────────────────

const children = [];

// Title / Subtitle
children.push(titlePara("SNOP Platform v2 — Complete Backend Python Code"));
children.push(subtitlePara("FastAPI + PostgreSQL + SQLAlchemy  |  Auto-generated"));

// ── SECTION 1 ────────────────────────────────────────────────────────────────
children.push(heading1("SECTION 1 — Application Entry Point"));
children.push(heading2("backend/app/main.py"));
children.push(...codeBlock(MAIN_PY));
children.push(spacer(200));

// ── SECTION 2 ────────────────────────────────────────────────────────────────
children.push(heading1("SECTION 2 — Core Configuration"));
children.push(heading2("backend/app/core/config.py"));
children.push(...codeBlock(CONFIG_PY));
children.push(spacer(200));
children.push(heading2("backend/app/core/database.py"));
children.push(...codeBlock(DATABASE_PY));
children.push(spacer(200));
children.push(heading2("backend/app/core/celery_app.py"));
children.push(...codeBlock(CELERY_PY));
children.push(spacer(200));

// ── SECTION 3 ────────────────────────────────────────────────────────────────
children.push(heading1("SECTION 3 — Database Models"));
children.push(heading2("backend/app/models/models.py"));
children.push(...codeBlock(MODELS_PY));
children.push(spacer(200));

// ── SECTION 4 ────────────────────────────────────────────────────────────────
children.push(heading1("SECTION 4 — Pydantic Schemas"));
children.push(heading2("backend/app/schemas/schemas.py"));
children.push(...codeBlock(SCHEMAS_PY));
children.push(spacer(200));

// ── SECTION 5 ────────────────────────────────────────────────────────────────
children.push(heading1("SECTION 5 — API Routes"));
const routes = [
  ["backend/app/api/routes/health.py",        HEALTH_PY],
  ["backend/app/api/routes/dashboard.py",     DASHBOARD_PY],
  ["backend/app/api/routes/plants.py",        PLANTS_PY],
  ["backend/app/api/routes/furnaces.py",      FURNACES_PY],
  ["backend/app/api/routes/feeds.py",         FEEDS_PY],
  ["backend/app/api/routes/yield_matrix.py",  YIELD_MATRIX_PY],
  ["backend/app/api/routes/scenarios.py",     SCENARIOS_PY],
  ["backend/app/api/routes/optimization.py",  OPTIMIZATION_PY],
  ["backend/app/api/routes/uploads.py",       UPLOADS_PY],
  ["backend/app/api/routes/regression.py",    REGRESSION_PY],
  ["backend/app/api/routes/sensitivity.py",   SENSITIVITY_PY],
];
for (const [label, src] of routes) {
  children.push(heading2(label));
  children.push(...codeBlock(src));
  children.push(spacer(200));
}

// ── SECTION 6 ────────────────────────────────────────────────────────────────
children.push(heading1("SECTION 6 — Services"));
const services = [
  ["backend/app/services/excel_service.py",                    EXCEL_SVC_PY],
  ["backend/app/services/optimizer.py",                        OPTIMIZER_PY],
  ["backend/app/services/optimization_tasks.py",               OPT_TASKS_PY],
  ["backend/app/services/regression_service.py",               REG_SVC_PY],
  ["backend/app/services/sensitivity_regression_service.py",   SENS_REG_SVC_PY],
  ["backend/app/services/sensitivity_upload_service.py",       SENS_UPL_SVC_PY],
];
for (const [label, src] of services) {
  children.push(heading2(label));
  children.push(...codeBlock(src));
  children.push(spacer(200));
}

// ── SECTION 7 ────────────────────────────────────────────────────────────────
children.push(heading1("SECTION 7 — Input / Output File Formats"));

// 7.1 Sensitivity Upload
children.push(heading3("7.1  Sensitivity Data Upload Format  (POST /api/sensitivity/upload)"));
children.push(new Paragraph({
  spacing: { before: 80, after: 80 },
  children: [new TextRun({
    text: "This is the main upload for plant-level simulation data. No furnace_id column needed.",
    font: BODY_FONT, size: 20, color: "000000"
  })]
}));
children.push(new Paragraph({
  spacing: { before: 120, after: 60 },
  children: [new TextRun({ text: "Required columns (only COT is mandatory):", font: BODY_FONT, size: 20, bold: true })]
}));
children.push(makeTable(
  ["Column", "Accepted Names", "Type", "Description"],
  [
    ["COT",          "cot, coil_outlet_temp, cot_celsius, cot_actual",                           "Float", "Coil Outlet Temperature (deg C) — REQUIRED"],
    ["CIP",          "cip, cracking_intensity, cracking_intensity_parameter, cip_value",          "Float", "Cracking Intensity Parameter"],
    ["Feed Rate",    "feed_rate, feedrate, feed rate, feed",                                      "Float", "Feed rate (t/h)"],
    ["SHC",          "shc, steam_hc_ratio, dilution_steam, dsr",                                 "Float", "Steam/HC ratio"],
    ["CIT",          "cit, coil_inlet_temp, inlet_temp",                                         "Float", "Coil Inlet Temperature (deg C)"],
    ["COP",          "cop, coil_outlet_pressure, outlet_pressure",                               "Float", "Coil Outlet Pressure (bar)"],
    ["Feed Ethane",  "feed_ethane, ethane%, ethane_pct",                                         "Float", "Ethane fraction in feed (%)"],
    ["Feed Propane", "feed_propane, propane%, propane_in",                                       "Float", "Propane fraction in feed (%)"],
  ],
  [1600, 3400, 900, 3400]
));
children.push(new Paragraph({
  spacing: { before: 120, after: 60 },
  children: [new TextRun({ text: "Output columns (all optional):", font: BODY_FONT, size: 20, bold: true })]
}));
children.push(makeTable(
  ["Column", "Accepted Names", "Description"],
  [
    ["H2 Yield",      "hydrogen_yield, h2, h2_yield",                            "Hydrogen yield"],
    ["CH4 Yield",     "methane_yield, ch4, ch4_yield",                           "Methane yield"],
    ["Coking Rate",   "coking_rate, coke_rate, coking",                          "Coking rate"],
    ["Heat Absorbed", "heat_absorbed, heat absorbed per radiant coil",            "Heat absorbed per radiant coil"],
    ["C2H4 Yield",    "ethylene_yield, c2h4_yield, ethylene",                    "Ethylene yield"],
    ["C2H6 Yield",    "c2h6_yield, c2h6, ethane_yield",                         "Ethane yield"],
    ["Benzene Exit",  "benzene_exit, benzene_coil_exit, benzene",                "Benzene coil exit"],
    ["Styrene Exit",  "styrene_exit, styrene_coil_exit, styrene",                "Styrene coil exit"],
    ["C2H2 Exit",     "c2h2_exit, c2h2_coil_exit, acetylene",                   "Acetylene coil exit"],
  ],
  [1600, 3800, 3900]
));
children.push(new Paragraph({ spacing: { before: 120, after: 60 }, children: [new TextRun({ text: "Notes:", font: BODY_FONT, size: 20, bold: true })] }));
const notes71 = [
  "File can be Excel (.xlsx/.xlsm/.xls) or CSV",
  "Multi-sheet Excel: each sheet name is used as furnace_id automatically",
  'Single-sheet without furnace_id column: data stored as furnace_id="ALL" (plant-level, applies to all furnaces H1-H6)',
  "Column names are case-insensitive and space/underscore variants are accepted",
];
for (const n of notes71) {
  children.push(new Paragraph({
    spacing: { before: 40, after: 40 },
    indent: { left: 360 },
    children: [new TextRun({ text: "- " + n, font: BODY_FONT, size: 20, color: "000000" })]
  }));
}

// 7.2 Yield Upload
children.push(heading3("7.2  Yield Data Upload Format  (POST /api/uploads/, upload_type=yield_data)"));
children.push(makeTable(
  ["Column", "Note"],
  [
    ["furnace (code)", "Required"],
    ["feed (feedstock code)", "Required"],
    ["date", "Required"],
    ["ethylene, propylene, butadiene, benzene, toluene, hydrogen, methane, fuel_oil, pygas, c4, cot", "Optional — yield percentages"],
  ],
  [3000, 6300]
));

// 7.3 Feed Prices
children.push(heading3("7.3  Feed Prices Upload  (upload_type=feed_prices)"));
children.push(makeTable(
  ["Column", "Note"],
  [
    ["feed (feedstock code)", "Required"],
    ["date", "Required"],
    ["price", "Required — USD per tonne"],
  ],
  [3000, 6300]
));

// 7.4 Product Prices
children.push(heading3("7.4  Product Prices Upload  (upload_type=product_prices)"));
children.push(makeTable(
  ["Column", "Note"],
  [
    ["product", "Required"],
    ["date", "Required"],
    ["price", "Required — USD per tonne"],
  ],
  [3000, 6300]
));

// 7.5 Operating Data
children.push(heading3("7.5  Operating Data Upload  (upload_type=operating_data)"));
children.push(makeTable(
  ["Column", "Note"],
  [
    ["furnace (code)", "Required"],
    ["date", "Required"],
    ["feed_rate", "Required — t/h"],
    ["feedstock, cot, dilution, runtime, onstream", "Optional"],
  ],
  [3000, 6300]
));

// 7.6 Yield Matrix API Response
children.push(heading3("7.6  Yield Matrix API Response  (POST /api/sensitivity/yield-matrix)"));
children.push(new Paragraph({
  spacing: { before: 80, after: 40 },
  children: [new TextRun({ text: "Request body:", font: BODY_FONT, size: 20, bold: true })]
}));
const reqJson = `{
  "plant_code": "PL1",
  "furnace_ids": ["H-1","H-2"],
  "conditions": {
    "cot": 845,
    "cip": 0.65,
    "feed_rate": 26.5,
    "shc": 0.35,
    "cit": 620,
    "cop": 1.8,
    "feed_ethane": 5.0,
    "feed_propane": 2.0
  }
}`;
children.push(...codeBlock(reqJson));
children.push(new Paragraph({
  spacing: { before: 120, after: 40 },
  children: [new TextRun({ text: "Response:", font: BODY_FONT, size: 20, bold: true })]
}));
const respJson = `{
  "furnaces": ["H-1","H-2","H-3","H-4","H-5","H-6"],
  "furnace_meta": [{"code":"H-1","name":"Furnace H-1"},],
  "products": ["hydrogen_yield","methane_yield","coking_rate","heat_absorbed",
               "ethylene_yield","c2h6_yield","benzene_exit","styrene_exit","c2h2_exit"],
  "product_labels": {"hydrogen_yield":"H2","ethylene_yield":"C2H4 Yield",},
  "matrix": {
    "H-1": {"hydrogen_yield":2.1234,"ethylene_yield":32.5678,},
    "H-2": {}
  },
  "model_info": {"H-1":{"hydrogen_yield":{"model_type":"linear","r2":0.98,"source":"plant_level"}}},
  "conditions_used": {"cot":845.0,"cip":0.65,},
  "data_source": "plant_level"
}`;
children.push(...codeBlock(respJson));

// ── Assemble and write ────────────────────────────────────────────────────────

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: BODY_FONT, size: 20, color: "000000" } },
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run:       { size: 36, bold: true, font: BODY_FONT, color: DARK_BLUE },
        paragraph: { spacing: { before: 480, after: 160 }, outlineLevel: 0 },
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run:       { size: 22, bold: true, font: CODE_FONT, color: DARK_GREY },
        paragraph: { spacing: { before: 280, after: 80 }, outlineLevel: 1 },
      },
      {
        id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run:       { size: 24, bold: true, font: BODY_FONT, color: DARK_GREY },
        paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 2 },
      },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 },
      },
    },
    children,
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  const kb = Math.round(buf.length / 1024);
  console.log(`Written: ${OUT}  (${kb} KB)`);
}).catch(err => { console.error(err); process.exit(1); });
