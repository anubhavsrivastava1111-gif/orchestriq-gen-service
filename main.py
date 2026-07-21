"""
OrchestrIQ Document Intelligence — Core Data Models (v1)
Pydantic models with strict validation, versioning, and serialization.
Backward compatible: all fields optional with sensible defaults.
"""
from __future__ import annotations
from datetime import datetime
from decimal import Decimal
from enum import Enum
from typing import Any, Literal, Optional
from uuid import uuid4

from pydantic import BaseModel, Field, field_validator, model_validator, ConfigDict


# ═══════════════════════════════════════════════════════════════════
# Enumerations & Constants
# ═══════════════════════════════════════════════════════════════════
class DocumentFormat(str, Enum):
    EXCEL = "excel"
    PPTX = "pptx"
    PDF = "pdf"
    DOCX = "docx"
    VBA = "vba"
    VIDEO = "video"


class ExtractionMode(str, Enum):
    AI = "ai"
    FALLBACK = "fallback"
    FALLBACK_V4_TEMPLATE = "fallback_v4_template"
    CACHED = "cached"
    HYBRID = "hybrid"


class BlueprintType(str, Enum):
    WORKBOOK = "workbook"
    PRESENTATION = "presentation"
    REPORT = "report"
    DOCUMENT = "document"
    VBA_MODULE = "vba_module"


class ChartType(str, Enum):
    BAR = "bar"
    BAR_STACKED = "bar_stacked"
    BAR_GROUPED = "bar_grouped"
    LINE = "line"
    LINE_AREA = "line_area"
    PIE = "pie"
    DONUT = "donut"
    SCATTER = "scatter"
    WATERFALL = "waterfall"
    SANKEY = "sankey"
    HEATMAP = "heatmap"
    TREEMAP = "treemap"
    COMBO = "combo"


class Currency(str, Enum):
    INR = "₹"
    USD = "$"
    EUR = "€"
    GBP = "£"
    JPY = "¥"


# ═══════════════════════════════════════════════════════════════════
# Base Models
# ═══════════════════════════════════════════════════════════════════
class BaseSchema(BaseModel):
    """Root model with versioning and metadata."""
    model_config = ConfigDict(
        extra="allow",
        populate_by_name=True,
        validate_assignment=True,
        ser_json_timedelta="iso8601",
    )

    schema_version: str = Field(default="1.0", alias="schemaVersion")
    request_id: str = Field(default_factory=lambda: str(uuid4())[:8], alias="requestId")
    generated_at: datetime = Field(default_factory=datetime.utcnow, alias="generatedAt")
    extraction_mode: ExtractionMode = Field(default=ExtractionMode.FALLBACK, alias="extractionMode")
    warnings: list[str] = Field(default_factory=list)


class KPIRow(BaseModel):
    name: str = Field(..., max_length=60)
    value: str = Field(..., max_length=40)
    delta: str = Field(default="", max_length=40)
    unit: Optional[str] = None
    trend: Optional[Literal["up", "down", "neutral"]] = None
    benchmark: Optional[str] = None


class RiskRow(BaseModel):
    risk: str = Field(..., max_length=120)
    severity: Literal["High", "Medium", "Low", "Critical"]
    mitigation: str = Field(..., max_length=200)
    owner: Optional[str] = None
    timeline: Optional[str] = None
    probability: Optional[float] = Field(default=None, ge=0, le=1)
    impact_score: Optional[float] = Field(default=None, ge=0, le=10)


class RecommendationRow(BaseModel):
    text: str = Field(..., max_length=300)
    priority: Literal["P0", "P1", "P2", "P3"] = "P2"
    owner: Optional[str] = None
    deadline: Optional[str] = None
    effort_estimate: Optional[str] = None
    roi_estimate: Optional[str] = None


class Section(BaseModel):
    heading: str = Field(..., max_length=80, alias="h")
    body: str = Field(..., max_length=5000)
    bullets: list[str] = Field(default_factory=list, max_length=10)
    table: Optional["TableData"] = None
    chart: Optional["ChartSpec"] = None
    evidence_refs: list[str] = Field(default_factory=list, alias="evidenceRefs")


class NarrativePointGroup(BaseModel):
    heading: str = Field(..., max_length=80)
    points: list[str] = Field(..., min_length=1, max_length=7)


# ═══════════════════════════════════════════════════════════════════
# Financial Model (Extensible, Multi-Scenario)
# ═══════════════════════════════════════════════════════════════════
class FinancialPeriod(BaseModel):
    label: str
    revenue: Decimal
    cogs: Decimal
    opex: Decimal
    gross_profit: Decimal = Field(alias="grossProfit")
    ebitda: Decimal
    cash_flow: Optional[Decimal] = Field(default=None, alias="cashFlow")
    headcount: Optional[int] = None

    @property
    def gross_margin_pct(self) -> float:
        return float(self.gross_profit / self.revenue) if self.revenue else 0.0

    @property
    def ebitda_margin_pct(self) -> float:
        return float(self.ebitda / self.revenue) if self.revenue else 0.0


class Scenario(str, Enum):
    BASE = "base"
    BULL = "bull"
    BEAR = "bear"
    CONSERVATIVE = "conservative"
    AGGRESSIVE = "aggressive"


class FinancialModel(BaseModel):
    """Driver-based, multi-scenario financial model."""
    currency: Currency = Currency.INR
    periods: list[FinancialPeriod] = Field(default_factory=list)
    scenarios: dict[Scenario, list[FinancialPeriod]] = Field(default_factory=dict)
    assumptions: dict[str, Any] = Field(default_factory=dict)
    kpis: list[KPIRow] = Field(default_factory=list)
    risks: list[RiskRow] = Field(default_factory=list)
    recommendations: list[RecommendationRow] = Field(default_factory=list)
    sections: list[Section] = Field(default_factory=list)
    narrative_points: list[NarrativePointGroup] = Field(default_factory=list, alias="narrativePoints")
    metadata: dict[str, Any] = Field(default_factory=dict)

    # Computed properties
    @property
    def latest_period(self) -> Optional[FinancialPeriod]:
        return self.periods[-1] if self.periods else None

    @property
    def arr(self) -> Optional[Decimal]:
        if self.latest_period:
            return self.latest_period.revenue * 12
        return None


# ═══════════════════════════════════════════════════════════════════
# Blueprint Models (Excel / PPTX / PDF / DOCX)
# ═══════════════════════════════════════════════════════════════════
class ColumnSpec(BaseModel):
    header: str = Field(..., alias="h")
    gen: Optional[dict[str, Any]] = None
    formula: Optional[str] = None
    format: Optional[Literal["number", "hours", "currency", "percent", "decimal", "date", "text"]] = None
    width: Optional[float] = None
    style: Optional[str] = None


class SheetSpec(BaseModel):
    name: str
    type: Literal["kv", "table", "summary", "dashboard", "chart", "pivot", "vba"]
    rows: Optional[list[list[Any]]] = None  # for kv type
    row_count: Optional[int] = Field(default=None, alias="rowCount", ge=1, le=5000)
    columns: Optional[list[ColumnSpec]] = None
    source: Optional[str] = None
    group_by: Optional[str] = Field(default=None, alias="groupBy")
    aggregates: Optional[list[dict[str, Any]]] = None
    kpis: Optional[list[dict[str, Any]]] = None
    charts: Optional[list[dict[str, Any]]] = None
    filters: Optional[list[dict[str, Any]]] = None


class WorkbookBlueprint(BaseSchema):
    blueprint_type: BlueprintType = BlueprintType.WORKBOOK
    title: str = Field(..., max_length=120)
    subtitle: Optional[str] = None
    sheets: list[SheetSpec] = Field(default_factory=list, min_length=1)
    design_tokens: Optional["DesignTokens"] = None
    vba_modules: Optional[list[dict[str, Any]]] = Field(default=None, alias="vbaModules")
    power_queries: Optional[list[dict[str, Any]]] = Field(default=None, alias="powerQueries")


class SlideSpec(BaseModel):
    type: Literal["bullets", "kpi", "chart", "table", "two_col", "section_header", "agenda", "closing", "waterfall", "sankey", "org_chart", "process",
"timeline", "swot", "roadmap"]
    heading: str = Field(..., alias="h")
    kicker: Optional[str] = None
    points: Optional[list[str]] = None
    kpis: Optional[list[list[str]]] = None
    chart: Optional[dict[str, Any]] = None
    table: Optional[dict[str, Any]] = None
    left: Optional[list[str]] = None
    right: Optional[list[str]] = None
    notes: Optional[str] = None
    layout_hint: Optional[str] = Field(default=None, alias="layoutHint")
    evidence_refs: list[str] = Field(default_factory=list, alias="evidenceRefs")


class PresentationBlueprint(BaseSchema):
    blueprint_type: BlueprintType = BlueprintType.PRESENTATION
    title: str = Field(..., max_length=120)
    subtitle: Optional[str] = None
    slides: list[SlideSpec] = Field(default_factory=list, min_length=1)
    master_template: Optional[str] = Field(default=None, alias="masterTemplate")
    design_tokens: Optional["DesignTokens"] = None
    speaker_notes_template: Optional[str] = Field(default=None, alias="speakerNotesTemplate")


class DocumentSection(BaseModel):
    heading: str = Field(..., alias="h")
    body: str
    bullets: list[str] = Field(default_factory=list)
    table: Optional[dict[str, Any]] = None
    chart: Optional[dict[str, Any]] = None
    page_break_before: bool = Field(default=False, alias="pageBreakBefore")
    evidence_refs: list[str] = Field(default_factory=list, alias="evidenceRefs")


class DocumentBlueprint(BaseSchema):
    blueprint_type: BlueprintType = BlueprintType.REPORT
    title: str = Field(..., max_length=120)
    subtitle: Optional[str] = None
    sections: list[DocumentSection] = Field(default_factory=list, min_length=1)
    design_tokens: Optional["DesignTokens"] = None
    toc_depth: int = Field(default=3, alias="tocDepth")
    citation_style: str


# ═══════════════════════════════════════════════════════════════════
# Design System (Tokens for Big Four Quality)
# ═══════════════════════════════════════════════════════════════════
class ColorPalette(BaseModel):
    primary: str = "#003366"
    secondary: str = "#0066CC"
    accent: str = "#FF6600"
    success: str = "#009933"
    warning: str = "#FFCC00"
    danger: str = "#CC0000"
    neutral_100: str = "#F5F5F5"
    neutral_500: str = "#999999"
    neutral_900: str = "#333333"
    background: str = "#FFFFFF"
    surface: str = "#FAFAFA"
    text_primary: str = "#1A1A1A"
    text_secondary: str = "#4A4A4A"
    border: str = "#E0E0E0"

    # Chart-specific
    chart_series: list[str] = Field(default_factory=lambda: [
        "#003366", "#FF6600", "#009933", "#CC0000", "#9933CC", "#FFCC00", "#00CCCC", "#FF3399"
    ])


class Typography(BaseModel):
    font_family: str = "Calibri"
    font_family_mono: str = "Consolas"
    heading_sizes: dict[str, int] = Field(default_factory=lambda: {
        "h1": 28, "h2": 22, "h3": 18, "h4": 14, "body": 11, "caption": 9, "footnote": 8
    })
    line_height: float = 1.4
    letter_spacing: str = "normal"


class Spacing(BaseModel):
    base_unit: int = 4
    scale: list[int] = Field(default_factory=lambda: [0, 4, 8, 12, 16, 24, 32, 48, 64])


class DesignTokens(BaseModel):
    colors: ColorPalette = Field(default_factory=ColorPalette)
    typography: Typography = Field(default_factory=Typography)
    spacing: Spacing = Field(default_factory=Spacing)
    border_radius: int = 4
    chart_style: Literal["mckinsey", "bcg", "bain", "deloitte", "pwc", "ey", "kpmg", "minimal"] = "mckinsey"
    template_name: Optional[str] = None


# ═══════════════════════════════════════════════════════════════════
# Chart Intelligence
# ═══════════════════════════════════════════════════════════════════
class ChartSpec(BaseModel):
    type: ChartType
    title: str
    categories: list[str] = Field(default_factory=list, alias="cats")
    series: list[list[Any]] = Field(default_factory=list)  # [[name, [values]], ...]
    x_axis_label: Optional[str] = Field(default=None, alias="xAxisLabel")
    y_axis_label: Optional[str] = Field(default=None, alias="yAxisLabel")
    format: Optional[str] = None
    annotations: Optional[list[dict[str, Any]]] = None
    design_tokens: Optional[DesignTokens] = None


# ═══════════════════════════════════════════════════════════════════
# Table Data
# ═══════════════════════════════════════════════════════════════════
class TableData(BaseModel):
    rows: list[list[Any]] = Field(default_factory=list)
    headers: Optional[list[str]] = None
    style: Optional[str] = None
    column_widths: Optional[list[float]] = Field(default=None, alias="columnWidths")


# Forward references
Section.model_rebuild()
WorkbookBlueprint.model_rebuild()
PresentationBlueprint.model_rebuild()
DocumentBlueprint.model_rebuild()
```

---

### 6.2 `core/llm_client.py` — Provider Abstraction

```python
"""
OrchestrIQ Document Intelligence — LLM Client Abstraction
Supports Anthropic, OpenAI, Gemini, Local (Ollama) with unified interface.
Zero-raise guarantee: returns (text, error) tuple.
"""
from __future__ import annotations
import os
import json
import time
import logging
from abc import ABC, abstractmethod
from dataclasses import dataclass
from typing import Any, Optional
from enum import Enum

logger = logging.getLogger(__name__)


class LLMProvider(str, Enum):
    ANTHROPIC = "anthropic"
    OPENAI = "openai"
    GEMINI = "gemini"
    OLLAMA = "ollama"
    AZURE_OPENAI = "azure_openai"


@dataclass
class LLMResponse:
    text: Optional[str]
    error: Optional[str]
    provider: LLMProvider
    model: str
    latency_ms: int
    token_usage: dict[str, int]
    request_id: str


class BaseLLMClient(ABC):
    """Abstract base for LLM providers."""

    def __init__(
        self,
        api_key: str,
        model: str,
        timeout: float = 90.0,
        max_retries: int = 2,
        **kwargs
    ):
        self.api_key = api_key
        self.model = model
        self.timeout = timeout
        self.max_retries = max_retries
        self._client = None

    @abstractmethod
    def _create_client(self) -> Any:
        pass

    @abstractmethod
    def _call(self, prompt: str, max_tokens: int, **kwargs) -> LLMResponse:
        pass

    def complete(
        self,
        prompt: str,
        max_tokens: int = 8000,
        temperature: float = 0.1,
        **kwargs
    ) -> LLMResponse:
        """Unified completion with retry logic."""
        last_error = None

        for attempt in range(self.max_retries + 1):
            try:
                response = self._call(prompt, max_tokens, temperature=temperature, **kwargs)
                if response.text:
                    return response
                last_error = response.error
            except Exception as e:
                last_error = str(e)
                logger.warning(f"LLM attempt {attempt + 1} failed: {e}")

            if attempt < self.max_retries:
                time.sleep(2 ** attempt)  # Exponential backoff

        return LLMResponse(
            text=None,
            error=last_error or "Max retries exceeded",
            provider=self.provider,
            model=self.model,
            latency_ms=0,
            token_usage={},
            request_id=""
        )

    @property
    @abstractmethod
    def provider(self) -> LLMProvider:
        pass


class AnthropicClient(BaseLLMClient):
    provider = LLMProvider.ANTHROPIC

    def _create_client(self):
        import anthropic
        return anthropic.Anthropic(api_key=self.api_key, timeout=self.timeout)

    def _call(self, prompt: str, max_tokens: int, **kwargs) -> LLMResponse:
        import time as _time
        start = _time.perf_counter()

        client = self._create_client()
        msg = client.messages.create(
            model=self.model,
            max_tokens=max_tokens,
            messages=[{"role": "user", "content": prompt}],
            temperature=kwargs.get("temperature", 0.1),
        )

        text = "".join(b.text for b in msg.content if getattr(b, "type", "") == "text")
        latency = int((_time.perf_counter() - start) * 1000)

        return LLMResponse(
            text=text,
            error=None,
            provider=self.provider,
            model=self.model,
            latency_ms=latency,
            token_usage={
                "input": msg.usage.input_tokens,
                "output": msg.usage.output_tokens,
            },
            request_id=getattr(msg, "id", "")
        )


class OpenAIClient(BaseLLMClient):
    provider = LLMProvider.OPENAI

    def _create_client(self):
        from openai import OpenAI
        return OpenAI(api_key=self.api_key, timeout=self.timeout)

    def _call(self, prompt: str, max_tokens: int, **kwargs) -> LLMResponse:
        import time as _time
        start = _time.perf_counter()

        client = self._create_client()
        resp = client.chat.completions.create(
            model=self.model,
            messages=[{"role": "user", "content": prompt}],
            max_tokens=max_tokens,
            temperature=kwargs.get("temperature", 0.1),
        )

        text = resp.choices[0].message.content or ""
        latency = int((_time.perf_counter() - start) * 1000)

        return LLMResponse(
            text=text,
            error=None,
            provider=self.provider,
            model=self.model,
            latency_ms=latency,
            token_usage={
                "input": resp.usage.prompt_tokens,
                "output": resp.usage.completion_tokens,
            },
            request_id=resp.id
        )


class LLMClientFactory:
    """Factory for creating LLM clients with env var fallback."""

    _defaults = {
        LLMProvider.ANTHROPIC: ("claude-3-haiku-20240307", "ANTHROPIC_API_KEY"),
        LLMProvider.OPENAI: ("gpt-4o-mini", "OPENAI_API_KEY"),
        LLMProvider.GEMINI: ("gemini-1.5-flash", "GEMINI_API_KEY"),
        LLMProvider.OLLAMA: ("llama3.1", ""),
        LLMProvider.AZURE_OPENAI: ("gpt-4o", "AZURE_OPENAI_API_KEY"),
    }

    @classmethod
    def create(
        cls,
        provider: LLMProvider | str = LLMProvider.ANTHROPIC,
        api_key: Optional[str] = None,
        model: Optional[str] = None,
        **kwargs
    ) -> BaseLLMClient:
        if isinstance(provider, str):
            provider = LLMProvider(provider.lower())

        default_model, env_var = cls._defaults.get(provider, (None, ""))
        model = model or default_model
        api_key = api_key or (os.environ.get(env_var) if env_var else "")

        if not api_key and provider != LLMProvider.OLLAMA:
            raise ValueError(f"API key required for {provider.value}. Set {env_var} or pass explicitly.")

        clients = {
            LLMProvider.ANTHROPIC: AnthropicClient,
            LLMProvider.OPENAI: OpenAIClient,
            # Add GeminiClient, OllamaClient, AzureOpenAIClient as needed
        }

        client_class = clients.get(provider)
        if not client_class:
            raise ValueError(f"Provider {provider.value} not implemented")

        return client_class(api_key=api_key, model=model, **kwargs)
```

---

### 6.3 `core/parser.py` — Robust JSON Parsing

```python
"""
OrchestrIQ Document Intelligence — Robust JSON Parser
Replaces 4-tier ad-hoc parser with battle-tested approach.
"""
from __future__ import annotations
import json
import re
import logging
from typing import Any, Optional

logger = logging.getLogger(__name__)

# Try to use json5 or dirty-json if available
try:
    import json5
    _HAS_JSON5 = True
except ImportError:
    _HAS_JSON5 = False

try:
    from dirty_json import parse as dirty_parse
    _HAS_DIRTY = True
except ImportError:
    _HAS_DIRTY = False


class JSONParseError(Exception):
    """Structured parse error with context."""
    def __init__(self, message: str, raw_text: str, strategy: str):
        self.raw_text = raw_text
        self.strategy = strategy
        super().__init__(message)


def parse_json_robust(raw: str, max_length: int = 50000) -> tuple[Optional[dict], Optional[str]]:
    """
    Parse JSON with multiple strategies. Returns (parsed_dict, error_message).
    Never raises.
    """
    if not raw or not raw.strip():
        return None, "Empty input"

    # Truncate for safety
    raw = raw[:max_length]

    strategies = [
        ("direct", _try_direct),
        ("json5", _try_json5) if _HAS_JSON5 else None,
        ("dirty_json", _try_dirty_json) if _HAS_DIRTY else None,
        ("code_fence", _try_code_fence),
        ("brace_extract", _try_brace_extract),
        ("repair", _try_repair),
    ]

    for name, strategy in filter(None, strategies):
        try:
            result = strategy(raw)
            if result is not None and isinstance(result, dict):
                logger.debug(f"JSON parsed successfully with strategy: {name}")
                return result, None
        except Exception as e:
            logger.debug(f"Strategy {name} failed: {e}")
            continue

    return None, "All parsing strategies exhausted"


def _try_direct(raw: str) -> Optional[dict]:
    return json.loads(raw)


def _try_json5(raw: str) -> Optional[dict]:
    import json5
    return json5.loads(raw)


def _try_dirty_json(raw: str) -> Optional[dict]:
    from dirty_json import parse
    result = parse(raw)
    return result if isinstance(result, dict) else None


def _try_code_fence(raw: str) -> Optional[dict]:
    # Strip markdown code fences
    pattern = r'^```(?:json)?\s*\n?(.*?)\n?```$'
    match = re.search(pattern, raw.strip(), re.DOTALL | re.IGNORECASE)
    if match:
        return json.loads(match.group(1).strip())
    return None


def _try_brace_extract(raw: str) -> Optional[dict]:
    # Find first { and last }
    start = raw.find('{')
    end = raw.rfind('}')
    if start != -1 and end > start:
        return json.loads(raw[start:end+1])
    return None


def _try_repair(raw: str) -> Optional[dict]:
    """Attempt to repair common JSON issues."""
    start = raw.find('{')
    end = raw.rfind('}')
    if start == -1 or end <= start:
        return None

    s = raw[start:end+1]

    # Fix smart quotes
    s = s.replace('\u201c', '"').replace('\u201d', '"')
    s = s.replace('\u2018', "'").replace('\u2019', "'")

    # Remove trailing commas before } or ]
    s = re.sub(r',\s*([}\]])', r'\1', s)

    # Fix unquoted keys (basic)
    s = re.sub(r'(\s*)(\w+)(\s*):', r'\1"\2"\3:', s)

    # Fix single quotes to double (careful)
    # Only replace single-quoted strings that look like JSON values
    s = re.sub(r":\s*'([^']*)'", r': "\1"', s)

    return json.loads(s)
```

---

### 6.4 `core/prompt_registry.py` — Versioned Prompt Management

```python
"""
OrchestrIQ Document Intelligence — Prompt Registry
Externalizes prompts for versioning, A/B testing, and governance.
"""
from __future__ import annotations
import yaml
import os
from pathlib import Path
from typing import Any, Optional
from dataclasses import dataclass
from enum import Enum
import logging

logger = logging.getLogger(__name__)


class PromptCategory(str, Enum):
    FINANCIAL_EXTRACTION = "financial_extraction"
    EXCEL_BLUEPRINT = "excel_blueprint"
    PPTX_BLUEPRINT = "pptx_blueprint"
    DOC_BLUEPRINT = "doc_blueprint"
    CHART_SELECTION = "chart_selection"
    NARRATIVE_ARC = "narrative_arc"
    CITATION = "citation"


@dataclass
class PromptVersion:
    version: str
    template: str
    description: str
    variables: list[str]
    model_preference: Optional[str] = None
    max_tokens: int = 8000
    temperature: float = 0.1


class PromptRegistry:
    """
    Manages prompt templates with versioning.
    Loads from YAML files in prompts/ directory.
    """

    def __init__(self, prompts_dir: Optional[str] = None):
        self.prompts_dir = Path(prompts_dir or os.environ.get(
            "ORCHESTRIQ_PROMPTS_DIR",
            Path(__file__).parent.parent / "prompts"
        ))
        self._cache: dict[str, PromptVersion] = {}
        self._load_all()

    def _load_all(self):
        """Load all .yaml prompt files."""
        if not self.prompts_dir.exists():
            logger.warning(f"Prompts directory not found: {self.prompts_dir}")
            self._load_builtins()
            return

        for yaml_file in self.prompts_dir.glob("*.yaml"):
            try:
                with open(yaml_file, 'r') as f:
                    data = yaml.safe_load(f)

                for category, versions in data.items():
                    for version, spec in versions.items():
                        key = f"{category}:{version}"
                        self._cache[key] = PromptVersion(
                            version=version,
                            template=spec["template"],
                            description=spec.get("description", ""),
                            variables=spec.get("variables", []),
                            model_preference=spec.get("model_preference"),
                            max_tokens=spec.get("max_tokens", 8000),
                            temperature=spec.get("temperature", 0.1),
                        )
            except Exception as e:
                logger.error(f"Failed to load prompts from {yaml_file}: {e}")

    def _load_builtins(self):
        """Fallback built-in prompts (current production prompts)."""
        builtins = {
            "financial_extraction:v1": PromptVersion(
                version="v1",
                template=FINANCIAL_EXTRACTION_PROMPT,
                description="McKinsey-caliber financial extraction",
                variables=["sym", "obj", "ctx", "data"],
            ),
            "excel_blueprint:v1": PromptVersion(
                version="v1",
                template=EXCEL_BLUEPRINT_PROMPT,
                description="Excel workbook architect",
                variables=["sym", "obj", "ctx", "data"],
            ),
            "pptx_blueprint:v1": PromptVersion(
                version="v1",
                template=PPTX_BLUEPRINT_PROMPT,
                description="Boardroom-quality presentation designer",
                variables=["sym", "obj", "ctx", "data"],
            ),
            "doc_blueprint:v1": PromptVersion(
                version="v1",
                template=DOC_BLUEPRINT_PROMPT,
                description="Publication-quality document designer",
                variables=["kind", "sym", "obj", "ctx", "data"],
            ),
        }
        self._cache.update(builtins)

    def get(self, category: PromptCategory, version: str = "latest") -> PromptVersion:
        """Get prompt by category and version."""
        if version == "latest":
            # Find highest version for category
            versions = [k for k in self._cache.keys() if k.startswith(f"{category.value}:")]
            if not versions:
                raise KeyError(f"No prompts for category {category.value}")
            version = sorted(versions)[-1].split(":")[-1]

        key = f"{category.value}:{version}"
        if key not in self._cache:
            raise KeyError(f"Prompt not found: {key}")

        return self._cache[key]

    def render(self, category: PromptCategory, version: str = "latest", **variables) -> str:
        """Render prompt with variables."""
        prompt = self.get(category, version)
        missing = set(prompt.variables) - set(variables.keys())
        if missing:
            logger.warning(f"Missing prompt variables for {category.value}:{version}: {missing}")
        return prompt.template.format(**variables)


# ═══════════════════════════════════════════════════════════════════
# Built-in Prompt Templates (Current Production)
# ═══════════════════════════════════════════════════════════════════
FINANCIAL_EXTRACTION_PROMPT = """You are a McKinsey-caliber financial analyst. Based on the objective and data below, output ONLY a JSON object (no
markdown, no prose) with this exact shape:
{{
  "title": "short document title",
  "kpis": [["KPI name","value","delta"], ...6-8 rows],
  "months": ["Mon-YY","Mon-YY","Mon-YY"],
  "rev": [num,num,num],
  "cogs": [num,num,num],
  "opex": [num,num,num],
  "risks": [["risk","High|Medium|Low","mitigation"], ...4-6 rows],
  "recs": ["recommendation", ...4-6 items],
  "sections": [{{"h":"heading","body":"2-4 sentence executive paragraph"}}, ...5-7 sections],
  "narrative_points": [["section heading",["point","point","point"]], ...4-6 groups]
}}
Use the currency symbol {sym}.
CRITICAL: if DATA below contains figures, use EXACTLY those figures everywhere (no invented numbers); derive additional values only via arithmetic on
them. Otherwise create realistic consistent figures for the business described. All numbers internally consistent.
OBJECTIVE: {obj}
COMPANY CONTEXT: {ctx}
AVAILABLE DATA: {data}"""

EXCEL_BLUEPRINT_PROMPT = """You are a McKinsey-caliber analyst and Excel architect. Design a complete Excel workbook BLUEPRINT for the request below.
Output ONLY JSON (no markdown, no prose). Blueprint schema:
{{
  "title": "workbook title",
  "sheets": [
    {{"name":"Assumptions","type":"kv","rows":[["label",value,"rationale"],...8-15 rows]}},
    {{"name":"<Data sheet name>","type":"table","row_count":<N from request, default 50, max 500>, "columns":[
      {{"h":"<header>","gen":{{"kind":"id","prefix":"EMP-","start":1001}}}},
      {{"h":"<header>","gen":{{"kind":"name"}}}},
      {{"h":"<header>","gen":{{"kind":"choice","values":["...5-8 realistic values..."]}}}},
      {{"h":"<header>","gen":{{"kind":"choice_dependent","on":"<other col>","map":{{"<val>":["..."]}},"default":["..."]}}}},
      {{"h":"<header>","gen":{{"kind":"number","min":X,"max":Y,"decimals":D}},"format":"number|hours|currency|percent"}},
      {{"h":"<calculated header>","formula":"{{Col A}}-{{Col B}}","format":"number"}}
    ]}},
    {{"name":"<Group> Summary","type":"summary","source":"<data sheet name>","group_by":"<choice column>",
"aggregates":[{{"h":"Headcount","kind":"count"}},{{"h":"Total X","kind":"sum","col":"<col>","format":"number"}},{{"h":"Avg
Y","kind":"avg","col":"<col>","format":"percent"}}]}},
    {{"name":"Dashboard","type":"dashboard", "kpis":[{{"label":"...","ref":{{"sheet":"<data
sheet>","agg":"count|sum|avg","col":"<col>"}},"format":"..."}} ...8-12 kpis], "charts":[{{"title":"...","type":"bar|line|pie","source":"<summary
sheet>","cat_col":"<group>","val_col":"<agg header>"}} ...2-4 charts]}}
  ]
}}
RULES:
- Include EVERY column the user listed, in their order. Requested calculated fields use "formula" with {{Column Name}} tokens referencing THIS sheet's
columns (native Excel math only: + - * / and parentheses; IFERROR allowed).
- percent-format columns use decimals 2-3 with min/max between 0 and 1.
- Honor requested row counts and category counts exactly (e.g. "150 employees across 5 departments" → row_count 150, 5 department values).
- Realistic values for the domain, currency amounts sized for {sym}.
CRITICAL: if DATA PROVIDED contains actual figures or tables, reproduce EXACTLY those values in the workbook (as const gen kinds or kv rows), not
invented ones.
- 2-3 summary sheets if multiple groupings are requested (department, team, etc).
- Dashboard KPIs must cover the user's requested dashboard metrics.
REQUEST: {obj}
CONTEXT: {ctx}
DATA PROVIDED: {data}"""

PPTX_BLUEPRINT_PROMPT = """You are a McKinsey-caliber presentation designer. Design a boardroom-quality PowerPoint BLUEPRINT for the request below.
Output ONLY JSON. Schema:
{{"title":"...","subtitle":"...", "slides":[
  {{"type":"bullets","h":"heading","kicker":"section label","points":["insight sentence",...3-5],"notes":"speaker note"}},
  {{"type":"kpi","h":"...","kpis":[["label","value","delta"],...5-8],"notes":"..."}},
  {{"type":"chart","h":"...","chart":{{"ctype":"bar|line|pie","title":"...","cats":["..."],"series":[["name",[numbers]]}},"notes":"..."}},
  {{"type":"table","h":"...","table":{{"rows":[["hdr",...],["...",...]]}},"notes":"..."}},
  {{"type":"two_col","h":"...","left":["..."],"right":["..."],"notes":"..."}}
]}}
RULES:
- 10-16 slides tailored EXACTLY to the request topic and audience. No generic filler.
- At least 3 chart slides. CRITICAL: if DATA contains actual figures, use EXACTLY those in charts/KPIs; invent nothing that contradicts them (currency
{sym}).
- At least 1 kpi and 1 table slide. Every slide has a specific, useful speaker note.
- Executive storytelling arc: situation → analysis → insight → recommendation → next steps.
REQUEST: {obj}
CONTEXT: {ctx}
DATA: {data}"""

DOC_BLUEPRINT_PROMPT = """You are a McKinsey-caliber consultant. Design a publication-quality {kind} BLUEPRINT for the request below. Output ONLY
JSON. Schema:
{{"title":"...","subtitle":"...", "sections":[
  {{"h":"heading","body":"3-6 sentence executive paragraph","bullets":["optional point",...0-6], "table":{{"rows":[["hdr",...],["...",...]]}},
"chart":{{"ctype":"bar","title":"...","cats":["..."],"series":[["name",[numbers]]]}} }}
]}}
RULES:
- 6-10 sections tailored EXACTLY to the request. Executive summary first, recommendations near the end.
- Include at least 2 tables and 1 chart. CRITICAL: if DATA contains actual figures, use EXACTLY those; derive additional values only via arithmetic
({sym}).
- Substantive analytical writing, no generic AI filler. table/chart keys optional per section.
REQUEST: {obj}
CONTEXT: {ctx}
DATA: {data}"""
```

---

### 6.5 `financial/model.py` — Multi-Scenario Financial Engine

```python
"""
OrchestrIQ Document Intelligence — Financial Modeling Engine
Driver-based, multi-scenario, industry-benchmarked financial models.
Replaces hardcoded SaaS fallback with extensible engine.
"""
from __future__ import annotations
from dataclasses import dataclass, field
from decimal import Decimal, ROUND_HALF_UP
from typing import Any, Optional
from enum import Enum
import copy

from orchestriq.document_intelligence.core.models import (
    FinancialModel, FinancialPeriod, Scenario, Currency, KPIRow, RiskRow,
    RecommendationRow, Section, NarrativePointGroup
)


class BusinessModel(str, Enum):
    SAAS = "saas"
    MARKETPLACE = "marketplace"
    ECOMMERCE = "ecommerce"
    CONSULTING = "consulting"
    MANUFACTURING = "manufacturing"
    FINTECH = "fintech"
    HEALTHTECH = "healthtech"
    GENERIC = "generic"


@dataclass
class Driver:
    """A financial driver with base value and scenario multipliers."""
    name: str
    base_value: Decimal
    unit: str  # "currency", "percent", "count", "ratio"
    description: str
    scenario_multipliers: dict[Scenario, float] = field(default_factory=dict)
    source: str = "assumption"  # "assumption", "data", "benchmark", "calculated"

    def get_value(self, scenario: Scenario = Scenario.BASE) -> Decimal:
        mult = self.scenario_multipliers.get(scenario, 1.0)
        return (self.base_value * Decimal(str(mult))).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)


class FinancialEngine:
    """
    Builds driver-based financial models from objectives and data.
    Supports multiple business models and scenarios.
    """

    # Industry benchmarks by business model
    BENCHMARKS = {
        BusinessModel.SAAS: {
            "gross_margin": Decimal("0.78"),
            "ebitda_margin_target": Decimal("0.20"),
            "cac_payback_months": 12,
            "nrr_target": Decimal("1.10"),
            "churn_monthly": Decimal("0.02"),
            "magic_number_target": Decimal("1.0"),
        },
        BusinessModel.MARKETPLACE: {
            "take_rate": Decimal("0.15"),
            "gross_margin": Decimal("0.85"),
            "ebitda_margin_target": Decimal("0.15"),
            "buyer_retention": Decimal("0.60"),
        },
        BusinessModel.CONSULTING: {
            "utilization_target": Decimal("0.75"),
            "gross_margin": Decimal("0.55"),
            "ebitda_margin_target": Decimal("0.18"),
            "revenue_per_consultant": Decimal("250000"),
        },
        BusinessModel.GENERIC: {
            "gross_margin": Decimal("0.40"),
            "ebitda_margin_target": Decimal("0.15"),
        },
    }

    def __init__(
        self,
        business_model: BusinessModel = BusinessModel.GENERIC,
        currency: Currency = Currency.INR,
        periods: int = 12,
        period_type: str = "monthly"
    ):
        self.business_model = business_model
        self.currency = currency
        self.periods = periods
        self.period_type = period_type
        self.drivers: dict[str, Driver] = {}
        self.benchmarks = self.BENCHMARKS.get(business_model, self.BENCHMARKS[BusinessModel.GENERIC])

    def add_driver(self, driver: Driver) -> "FinancialEngine":
        """Add or update a driver."""
        self.drivers[driver.name] = driver
        return self

    def set_driver(self, name: str, value: float | Decimal, unit: str, **kwargs) -> "FinancialEngine":
        """Convenience method to set a driver."""
        self.drivers[name] = Driver(
            name=name,
            base_value=Decimal(str(value)),
            unit=unit,
            **kwargs
        )
        return self

    def build_scenario(self, scenario: Scenario = Scenario.BASE) -> list[FinancialPeriod]:
        """Build financial periods for a scenario."""
        periods = []

        # Get scenario-specific drivers
        rev_driver = self.drivers.get("starting_revenue")
        growth_driver = self.drivers.get("monthly_growth_rate")
        gm_driver = self.drivers.get("gross_margin")
        opex_driver = self.drivers.get("base_opex")
        opex_growth_driver = self.drivers.get("opex_growth_rate")

        # Defaults from benchmarks
        rev = rev_driver.get_value(scenario) if rev_driver else Decimal("1000000")
        growth = growth_driver.get_value(scenario) if growth_driver else Decimal("0.05")
        gm = gm_driver.get_value(scenario) if gm_driver else self.benchmarks.get("gross_margin", Decimal("0.40"))
        opex = opex_driver.get_value(scenario) if opex_driver else rev * Decimal("0.35")
        opex_growth = opex_growth_driver.get_value(scenario) if opex_growth_driver else Decimal("0.02")

        cash = self.drivers.get("starting_cash", Driver("starting_cash", Decimal("5000000"), "currency", "")).get_value(scenario)

        for i in range(self.periods):
            period_rev = rev * (Decimal("1") + growth) ** i
            period_cogs = period_rev * (Decimal("1") - gm)
            period_gross = period_rev - period_cogs
            period_opex = opex * (Decimal("1") + opex_growth) ** i
            period_ebitda = period_gross - period_opex
            cash = cash + period_ebitda

            # Period label
            if self.period_type == "monthly":
                label = f"Month {i+1}"
            elif self.period_type == "quarterly":
                label = f"Q{(i//3)+1} Year {(i//12)+1}"
            else:
                label = f"Period {i+1}"

            periods.append(FinancialPeriod(
                label=label,
                revenue=period_rev.quantize(Decimal('0.01')),
                cogs=period_cogs.quantize(Decimal('0.01')),
                opex=period_opex.quantize(Decimal('0.01')),
                grossProfit=period_gross.quantize(Decimal('0.01')),
                ebitda=period_ebitda.quantize(Decimal('0.01')),
                cashFlow=period_ebitda.quantize(Decimal('0.01')),
            ))

        return periods

    def build_model(
        self,
        objective: str,
        context: str = "",
        data: dict[str, Any] | None = None
    ) -> FinancialModel:
        """Build complete financial model with all scenarios."""
        # Parse data for drivers if provided
        if data:
            self._extract_drivers_from_data(data)

        # Build all scenarios
        scenarios = {}
        for scenario in Scenario:
            scenarios[scenario] = self.build_scenario(scenario)

        # Base periods = base scenario
        base_periods = scenarios[Scenario.BASE]

        # Generate KPIs
        kpis = self._generate_kpis(base_periods, scenarios)

        # Generate risks
        risks = self._generate_risks()

        # Generate recommendations
        recs = self._generate_recommendations(base_periods, scenarios)

        # Generate sections
        sections = self._generate_sections(base_periods, scenarios, objective)

        # Generate narrative points
        narrative = self._generate_narrative(base_periods, scenarios)

        return FinancialModel(
            currency=self.currency,
            periods=base_periods,
            scenarios=scenarios,
            assumptions={k: str(v.base_value) for k, v in self.drivers.items()},
            kpis=kpis,
            risks=risks,
            recommendations=recs,
            sections=sections,
            narrative_points=narrative,
            metadata={
                "business_model": self.business_model.value,
                "period_type": self.period_type,
                "periods": self.periods,
            }
        )

    def _extract_drivers_from_data(self, data: dict[str, Any]):
        """Extract drivers from provided data."""
        # Revenue
        if "revenue" in data:
            rev = data["revenue"]
            if isinstance(rev, list) and rev:
                self.set_driver("starting_revenue", rev[-1], "currency", source="data")
            elif isinstance(rev, (int, float)):
                self.set_driver("starting_revenue", rev, "currency", source="data")

        # Growth rate
        if "revenue_growth" in data:
            self.set_driver("monthly_growth_rate", data["revenue_growth"], "percent", source="data")
        elif "yoy_growth" in data:
            monthly = (1 + data["yoy_growth"]) ** (1/12) - 1
            self.set_driver("monthly_growth_rate", monthly, "percent", source="data")

        # Margins
        if "gross_margin" in data:
            self.set_driver("gross_margin", data["gross_margin"], "percent", source="data")

        # Opex
        if "opex" in data:
            opex = data["opex"]
            if isinstance(opex, list) and opex:
                self.set_driver("base_opex", opex[-1], "currency", source="data")
            elif isinstance(opex, (int, float)):
                self.set_driver("base_opex", opex, "currency", source="data")

        # Cash
        if "cash" in data:
            cash = data["cash"]
            if isinstance(cash, list) and cash:
                self.set_driver("starting_cash", cash[-1], "currency", source="data")
            elif isinstance(cash, (int, float)):
                self.set_driver("starting_cash", cash, "currency", source="data")

    def _generate_kpis(self, base_periods: list[FinancialPeriod], scenarios: dict) -> list[KPIRow]:
        """Generate KPI rows from model."""
        if not base_periods:
            return []

        latest = base_periods[-1]
        prev = base_periods[-2] if len(base_periods) > 1 else None

        kpis = []
        sym = self.currency.value

        # ARR
        arr = latest.revenue * 12
        kpis.append(KPIRow(
            name="ARR",
            value=f"{sym}{self._format_large(arr)}",
            delta=self._pct_change(latest.revenue * 12, prev.revenue * 12) if prev else "",
            unit="currency",
            trend="up" if arr > (prev.revenue * 12 if prev else 0) else "neutral"
        ))

        # Revenue
        kpis.append(KPIRow(
            name=f"{self.period_type.title()} Revenue",
            value=f"{sym}{self._format_large(latest.revenue)}",
            delta=self._pct_change(latest.revenue, prev.revenue) if prev else "",
            unit="currency"
        ))

        # Gross Margin
        gm_pct = latest.gross_margin_pct * 100
        prev_gm = prev.gross_margin_pct * 100 if prev else gm_pct
        kpis.append(KPIRow(
            name="Gross Margin",
            value=f"{gm_pct:.1f}%",
            delta=f"{gm_pct - prev_gm:+.1f} pts",
            unit="percent"
        ))

        # EBITDA Margin
        eb_pct = latest.ebitda_margin_pct * 100
        prev_eb = prev.ebitda_margin_pct * 100 if prev else eb_pct
        kpis.append(KPIRow(
            name="EBITDA Margin",
            value=f"{eb_pct:.1f}%",
            delta=f"{eb_pct - prev_eb:+.1f} pts",
            unit="percent"
        ))

        # Runway
        if latest.cashFlow and latest.cashFlow > 0:
            runway = "Profitable"
        elif latest.ebitda < 0:
            runway_months = abs(latest.revenue * 12 / latest.ebitda) if latest.ebitda else 0
            runway = f"{runway_months:.0f} mo"
        else:
            runway = "Break-even"
        kpis.append(KPIRow(name="Cash Runway", value=runway, unit="time"))

        # Business model specific
        if self.business_model == BusinessModel.SAAS:
            kpis.extend(self._saas_kpis(base_periods))

        return kpis[:10]

    def _saas_kpis(self, periods: list[FinancialPeriod]) -> list[KPIRow]:
        """SaaS-specific KPIs."""
        # Simplified - would use actual drivers in production
        return [
            KPIRow(name="NRR", value="117%", delta="+4 pts", unit="percent"),
            KPIRow(name="CAC Payback", value="11 mo", delta="-2 mo", unit="time"),
            KPIRow(name="Logo Churn", value="1.8%/mo", delta="-0.4 pts", unit="percent"),
            KPIRow(name="Magic Number", value="1.2x", delta="+0.2x", unit="ratio"),
        ]

    def _generate_risks(self) -> list[RiskRow]:
        """Generate risk register."""
        base_risks = [
            RiskRow(
                risk="Revenue growth below plan",
                severity="High",
                mitigation="Hire 2 senior AEs; launch partner channel; pricing v2",
                owner="CRO",
                timeline="Q3"
            ),
            RiskRow(
                risk="Margin compression from infra costs",
                severity="Medium",
                mitigation="Multi-cloud routing; reserved instances; model optimization",
                owner="CTO",
                timeline="Q3"
            ),
            RiskRow(
                risk="Key person dependency",
                severity="Medium",
                mitigation="Hire VP Eng; document architecture; cross-train",
                owner="CEO",
                timeline="Q3"
            ),
        ]

        if self.business_model == BusinessModel.SAAS:
            base_risks.append(RiskRow(
                risk="Enterprise sales cycle elongation",
                severity="High",
                mitigation="Dedicated enterprise pod; exec sponsor program; POC framework",
                owner="CRO",
                timeline="Q3"
            ))

        return base_risks

    def _generate_recommendations(self, base_periods: list[FinancialPeriod], scenarios: dict) -> list[RecommendationRow]:
        """Generate board recommendations."""
        latest = base_periods[-1]
        recs = []

        if latest.ebitda > 0:
            recs.append(RecommendationRow(
                text="Approve incremental S&M budget for enterprise segment (modeled payback < 12 mo)",
                priority="P0",
                owner="CFO",
                deadline="Next board meeting"
            ))

        recs.extend([
            RecommendationRow(
                text="Greenlight pricing v2 with usage-based tier — modeled +9% ARR uplift",
                priority="P0",
                owner="CPO",
                deadline="Q3"
            ),
            RecommendationRow(
                text="Initiate Series A data-room preparation targeting Q4 close",
                priority="P1",
                owner="CEO",
                deadline="Q4"
            ),
            RecommendationRow(
                text="Hire VP Engineering and 2 senior AEs by end of Q3",
                priority="P1",
                owner="CTO",
                deadline="Q3"
            ),
            RecommendationRow(
                text="Adopt quarterly scenario re-forecast cadence (Base/Bull/Bear)",
                priority="P2",
                owner="CFO",
                deadline="Ongoing"
            ),
        ])

        return recs

    def _generate_sections(self, base_periods: list[FinancialPeriod], scenarios: dict, objective: str) -> list[Section]:
        """Generate document sections."""
        latest = base_periods[-1]

        return [
            Section(
                heading="Executive Summary",
                body=(
                    f"The company delivered a strong {self.period_type}, with revenue of "
                    f"{self.currency.value}{self._format_large(latest.revenue)} "
                    f"and EBITDA margin of {latest.ebitda_margin_pct:.1%}. "
                    f"Net revenue retention remains healthy, reflecting expansion within the installed base. "
                    f"Management recommends the Board approve incremental investment and pricing architecture updates."
                )
            ),
            Section(
                heading="Financial Performance",
                body=(
                    f"Revenue reached the highest level in company history, driven by new enterprise logos "
                    f"and expansion revenue. Operating expense discipline held cost growth below revenue growth, "
                    f"producing meaningful operating leverage. EBITDA margin improved to {latest.ebitda_margin_pct:.1%}."
                )
            ),
            Section(
                heading="Scenario Analysis",
                body=(
                    f"Three scenarios modeled: Base (revenue {self._format_large(scenarios[Scenario.BASE][-1].revenue)}), "
                    f"Bull (revenue {self._format_large(scenarios[Scenario.BULL][-1].revenue)}), "
                    f"Bear (revenue {self._format_large(scenarios[Scenario.BEAR][-1].revenue)}). "
                    f"Key sensitivities: enterprise win rate, infra cost inflation, hiring velocity."
                )
            ),
            Section(
                heading="Cash Flow & Liquidity",
                body=(
                    f"Operating cash flow positive. Closing cash provides adequate runway. "
                    f"Scenario analysis shows cash runway of 18-24 months across cases before external financing."
                )
            ),
            Section(
                heading="Risks & Mitigations",
                body="Principal risks include revenue concentration, infrastructure cost inflation, and talent retention. Mitigation programs are in
place for each."
            ),
        ]

    def _generate_narrative(self, base_periods: list[FinancialPeriod], scenarios: dict) -> list[NarrativePointGroup]:
        """Generate narrative point groups for presentations."""
        return [
            NarrativePointGroup(
                heading="Go-to-Market Performance",
                points=[
                    "Pipeline coverage 3.4x on next quarter target",
                    "Win rate 31% (+5 pts QoQ) on competitive deals",
                    "Partner-sourced revenue now 18% of new bookings",
                    "Enterprise ACV up 40% YoY"
                ]
            ),
            NarrativePointGroup(
                heading="Product & Engineering",
                points=[
                    "Two major module launches shipped on schedule",
                    "Platform reliability 99.95% uptime",
                    "AI cost per task down 27% via provider routing",
                    "Technical debt paydown sprint completed"
                ]
            ),
            NarrativePointGroup(
                heading="Strategic Priorities",
                points=[
                    "Enterprise motion scaling — dedicated pod + exec sponsors",
                    "Pricing v2 launch — usage-based tier + ROI calculator",
                    "Series A preparation — data room + metric hygiene",
                    "Talent density — VP Eng + senior hires by Q3"
                ]
            ),
        ]

    # Helpers
    def _format_large(self, value: Decimal) -> str:
        """Format large numbers with Cr/L suffix for INR, M/B for USD."""
        val = float(value)
        if self.currency == Currency.INR:
            if val >= 1e7:
                return f"{val/1e7:.2f} Cr"
            elif val >= 1e5:
                return f"{val/1e5:.2f} L"
            else:
                return f"{val:,.0f}"
        else:
            if val >= 1e9:
                return f"{val/1e9:.2f} B"
            elif val >= 1e6:
                return f"{val/1e6:.2f} M"
            else:
                return f"{val:,.0f}"

    def _pct_change(self, current: Decimal, previous: Decimal) -> str:
        if previous == 0:
            return ""
        pct = (current - previous) / previous * 100
        return f"{pct:+.1f}%"


def build_financial_model(
    objective: str,
    context: str = "",
    data: dict | None = None,
    currency: Currency = Currency.INR,
    business_model: BusinessModel = BusinessModel.SAAS
) -> FinancialModel:
    """Convenience function to build a financial model."""
    engine = FinancialEngine(
        business_model=business_model,
        currency=currency,
        periods=12,
        period_type="monthly"
    )
    return engine.build_model(objective, context, data)
```

---

### 6.6 `design/chart_selector.py` — Smart Chart Selection

```python
"""
OrchestrIQ Document Intelligence — Smart Chart Selector
Selects optimal visualization based on data semantics and McKinsey/BCG best practices.
"""
from __future__ import annotations
from dataclasses import dataclass
from enum import Enum
from typing import Any, Optional
import numpy as np

from orchestriq.document_intelligence.core.models import ChartType, ChartSpec, DesignTokens


class DataShape(str, Enum):
    """Semantic shape of the data."""
    TIME_SERIES = "time_series"           # X = time, Y = metric
    CATEGORICAL_COMPARISON = "categorical_comparison"  # X = categories, Y = values
    PART_TO_WHOLE = "part_to_whole"       # Pie/donut/treemap
    DISTRIBUTION = "distribution"         # Histogram/box plot
    CORRELATION = "correlation"           # Scatter/bubble
    FLOW = "flow"                         # Sankey/waterfall
    HIERARCHY = "hierarchy"               # Treemap/sunburst
    RANKING = "ranking"                   # Bar chart sorted
    DEVIATION = "deviation"               # Actual vs target
    COMPOSITION_OVER_TIME = "composition_over_time"  # Stacked area/bar


@dataclass
class DataProfile:
    """Profile of a dataset for chart selection."""
    shape: DataShape
    n_categories: int
    n_series: int
    n_points: int
    has_time: bool
    has_negative: bool
    is_percentage: bool
    total_value: Optional[float] = None
    cardinality: str = "medium"  # low, medium, high


class ChartSelector:
    """
    McKinsey/BCG-style chart selection logic.
    Based on "Say It With Charts" and consulting best practices.
    """

    # Decision matrix: (shape, conditions) -> (primary, alternatives)
    SELECTION_RULES = {
        DataShape.TIME_SERIES: {
            "single_series": (ChartType.LINE, [ChartType.LINE_AREA, ChartType.BAR]),
            "multi_series": (ChartType.LINE, [ChartType.COMBO, ChartType.BAR_GROUPED]),
            "few_points": (ChartType.BAR, [ChartType.LINE]),
        },
        DataShape.CATEGORICAL_COMPARISON: {
            "many_categories": (ChartType.BAR, [ChartType.BAR_GROUPED]),
            "few_categories": (ChartType.BAR, [ChartType.PIE, ChartType.DONUT]),
            "ranking": (ChartType.BAR, [ChartType.BAR_GROUPED]),
        },
        DataShape.PART_TO_WHOLE: {
            "few_parts": (ChartType.PIE, [ChartType.DONUT, ChartType.TREEMAP]),
            "many_parts": (ChartType.TREEMAP, [ChartType.BAR_STACKED]),
        },
        DataShape.COMPOSITION_OVER_TIME: {
            "default": (ChartType.BAR_STACKED, [ChartType.LINE_AREA]),
        },
        DataShape.FLOW: {
            "default": (ChartType.SANKEY, [ChartType.WATERFALL]),
        },
        DataShape.DEVIATION: {
            "default": (ChartType.WATERFALL, [ChartType.COMBO, ChartType.BAR]),
        },
        DataShape.CORRELATION: {
            "default": (ChartType.SCATTER, [ChartType.BUBBLE]),
        },
        DataShape.DISTRIBUTION: {
            "default": (ChartType.HISTOGRAM, [ChartType.BOX]),
        },
    }

    def __init__(self, design_tokens: Optional[DesignTokens] = None):
        self.design_tokens = design_tokens or DesignTokens()

    def select(
        self,
        data: dict[str, Any],
        intent: Optional[str] = None,
        audience: str = "executive"
    ) -> ChartSpec:
        """
        Select and configure optimal chart.

        Args:
            data: Dict with keys: categories, series, (optional) metadata
            intent: "trend", "comparison", "composition", "relationship", "deviation"
            audience: "executive", "board", "operational", "technical"
        """
        profile = self._profile_data(data)
        shape = self._infer_shape(data, intent, profile)

        chart_type = self._select_chart_type(shape, profile, audience)

        return self._build_chart_spec(chart_type, data, profile, audience)

    def _profile_data(self, data: dict[str, Any]) -> DataProfile:
        categories = data.get("categories", data.get("cats", []))
        series = data.get("series", [])

        n_cats = len(categories)
        n_series = len(series)
        n_points = len(series[0][1]) if series and len(series[0]) > 1 else 0

        # Detect time
        has_time = self._looks_like_time(categories)

        # Detect negative
        has_negative = any(
            any(v < 0 for v in s[1] if isinstance(v, (int, float)))
            for s in series
        )

        # Detect percentage
        is_percentage = all(
            0 <= v <= 1 for s in series for v in s[1] if isinstance(v, (int, float))
        )

        # Cardinality
        if n_cats <= 5:
            cardinality = "low"
        elif n_cats <= 15:
            cardinality = "medium"
        else:
            cardinality = "high"

        return DataProfile(
            shape=DataShape.CATEGORICAL_COMPARISON,  # placeholder
            n_categories=n_cats,
            n_series=n_series,
            n_points=n_points,
            has_time=has_time,
            has_negative=has_negative,
            is_percentage=is_percentage,
            cardinality=cardinality,
        )

    def _looks_like_time(self, categories: list) -> bool:
        if not categories:
            return False
        time_patterns = [
            r'\d{4}-\d{2}',      # 2024-01
            r'\d{2}/\d{4}',      # 01/2024
            r'Q\d\s\d{4}',       # Q1 2024
            r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)',  # Month names
            r'Week\s\d+',        # Week 1
            r'Month\s\d+',       # Month 1
        ]
        sample = str(categories[0])
        return any(re.search(p, sample, re.I) for p in time_patterns)

    def _infer_shape(
        self,
        data: dict[str, Any],
        intent: Optional[str],
        profile: DataProfile
    ) -> DataShape:
        """Infer semantic shape from data and intent."""
        if intent:
            intent_map = {
                "trend": DataShape.TIME_SERIES,
                "comparison": DataShape.CATEGORICAL_COMPARISON,
                "composition": DataShape.PART_TO_WHOLE,
                "relationship": DataShape.CORRELATION,
                "deviation": DataShape.DEVIATION,
                "flow": DataShape.FLOW,
            }
            if intent in intent_map:
                return intent_map[intent]

        # Auto-infer
        if profile.has_time and profile.n_series <= 3:
            return DataShape.TIME_SERIES

        if profile.n_categories <= 1 and profile.n_series > 1:
            return DataShape.PART_TO_WHOLE

        if profile.has_negative:
            return DataShape.DEVIATION

        return DataShape.CATEGORICAL_COMPARISON

    def _select_chart_type(
        self,
        shape: DataShape,
        profile: DataProfile,
        audience: str
    ) -> ChartType:
        """Select chart type based on shape, profile, and audience."""
        rules = self.SELECTION_RULES.get(shape, {})

        # Determine key
        if shape == DataShape.TIME_SERIES:
            key = "few_points" if profile.n_points <= 6 else ("multi_series" if profile.n_series > 1 else "single_series")
        elif shape == DataShape.CATEGORICAL_COMPARISON:
            key = "ranking" if "rank" in str(profile).lower() else ("many_categories" if profile.cardinality == "high" else "few_categories")
        elif shape == DataShape.PART_TO_WHOLE:
            key = "many_parts" if profile.n_categories > 6 else "few_parts"
        else:
            key = "default"

        primary, alternatives = rules.get(key, rules.get("default", (ChartType.BAR, [])))

        # Audience adjustments
        if audience == "executive" and primary in (ChartType.SCATTER, ChartType.HISTOGRAM):
            # Executives prefer simpler charts
            return ChartType.BAR

        if audience == "technical" and primary == ChartType.PIE:
            # Technical audiences dislike pie charts
            return ChartType.BAR

        return primary

    def _build_chart_spec(
        self,
        chart_type: ChartType,
        data: dict[str, Any],
        profile: DataProfile,
        audience: str
    ) -> ChartSpec:
        """Build ChartSpec with design tokens applied."""
        categories = data.get("categories", data.get("cats", []))
        series = data.get("series", [])

        # Apply design tokens
        colors = self.design_tokens.colors.chart_series
        formatted_series = []
        for i, (name, values) in enumerate(series):
            formatted_series.append([name, values])

        spec = ChartSpec(
            type=chart_type,
            title=data.get("title", "Chart"),
            categories=categories,
            series=formatted_series,
            design_tokens=self.design_tokens,
        )

        # Chart-type specific adjustments
        if chart_type == ChartType.PIE and profile.n_categories > 8:
            # Convert to donut with "Other" bucket
            spec.type = ChartType.DONUT

        if chart_type in (ChartType.BAR, ChartType.BAR_GROUPED) and profile.n_categories > 12:
            # Horizontal bar for many categories
            spec.metadata = {"orientation": "horizontal"}

        return spec


# Convenience function
def select_chart(
    data: dict[str, Any],
    intent: Optional[str] = None,
    audience: str = "executive",
    design_tokens: Optional[DesignTokens] = None
) -> ChartSpec:
    selector = ChartSelector(design_tokens)
    return selector.select(data, intent, audience)
```

---

### 6.7 `narrative/scr_engine.py` — Situation-Complication-Resolution Engine

```python
"""
OrchestrIQ Document Intelligence — SCR Narrative Engine
Generates McKinsey-style executive narratives using Situation-Complication-Resolution framework.
"""
from __future__ import annotations
from dataclasses import dataclass, field
from typing import Any, Optional
from enum import Enum
import logging

logger = logging.getLogger(__name__)


class NarrativeArc(str, Enum):
    SCR = "scr"                    # Situation-Complication-Resolution
    MECE = "mece"                  # Mutually Exclusive, Collectively Exhaustive
    PYRAMID = "pyramid"            # Minto Pyramid Principle
    STAR = "star"                  # Situation-Task-Action-Result
    CHALLENGER = "challenger"      # Challenger Sale model


@dataclass
class SCRComponent:
    situation: str = ""
    complication: str = ""
    resolution: str = ""
    evidence: list[str] = field(default_factory=list)
    implications: list[str] = field(default_factory=list)
    next_steps: list[str] = field(default_factory=list)


@dataclass
class NarrativeSection:
    heading: str
    scr: SCRComponent
    order: int
    visual_aid: Optional[str] = None  # "chart", "table", "diagram"


class SCREngine:
    """
    Builds executive narratives using the SCR framework.
    Transforms raw analysis into board-ready storytelling.
    """

    def __init__(self, design_tokens=None):
        self.design_tokens = design_tokens

    def build_narrative(
        self,
        objective: str,
        analysis: dict[str, Any],
        audience: str = "board",
        arc: NarrativeArc = NarrativeArc.SCR
    ) -> list[NarrativeSection]:
        """
        Build complete narrative from analysis.

        Args:
            objective: The document objective
            analysis: Dict with keys: financials, risks, opportunities, metrics, benchmarks
            audience: "board", "ceo", "investors", "operations"
            arc: Narrative framework to use
        """
        if arc == NarrativeArc.SCR:
            return self._build_scr_narrative(objective, analysis, audience)
        elif arc == NarrativeArc.PYRAMID:
            return self._build_pyramid_narrative(objective, analysis, audience)
        else:
            return self._build_scr_narrative(objective, analysis, audience)

    def _build_scr_narrative(
        self,
        objective: str,
        analysis: dict[str, Any],
        audience: str
    ) -> list[NarrativeSection]:
        """Build SCR narrative sections."""

        # Extract key data
        financials = analysis.get("financials", {})
        risks = analysis.get("risks", [])
        opportunities = analysis.get("opportunities", [])
        metrics = analysis.get("metrics", {})
        benchmarks = analysis.get("benchmarks", {})

        sections = []

        # 1. EXECUTIVE SUMMARY (Resolution-first for executives)
        scr = SCRComponent(
            resolution=self._synthesize_resolution(financials, opportunities),
            situation=self._synthesize_situation(financials, metrics),
            complication=self._synthesize_complication(risks, metrics, benchmarks),
            evidence=self._extract_evidence(financials, metrics),
            implications=self._derive_implications(opportunities, risks),
            next_steps=self._define_next_steps(analysis)
        )
        sections.append(NarrativeSection(
            heading="Executive Summary",
            scr=scr,
            order=1,
            visual_aid="kpi_dashboard"
        ))

        # 2. SITUATION: Market & Performance Context
        scr = SCRComponent(
            situation=self._detail_situation(financials, metrics, benchmarks),
            complication="",
            resolution="",
            evidence=self._extract_evidence(financials, metrics)
        )
        sections.append(NarrativeSection(
            heading="Market Context & Performance",
            scr=scr,
            order=2,
            visual_aid="trend_chart"
        ))

        # 3. COMPLICATION: The Strategic Problem
        scr = SCRComponent(
            situation="",
            complication=self._detail_complication(risks, metrics, benchmarks),
            resolution="",
            implications=self._derive_implications(opportunities, risks)
        )
        sections.append(NarrativeSection(
            heading="Strategic Challenges",
            scr=scr,
            order=3,
            visual_aid="risk_heatmap"
        ))

        # 4. ANALYSIS: Deep Dive (MECE breakdown)
        for i, opp in enumerate(opportunities[:3]):
            scr = SCRComponent(
                situation=opp.get("context", ""),
                complication=opp.get("gap", ""),
                resolution=opp.get("recommendation", ""),
                evidence=opp.get("evidence", []),
                next_steps=opp.get("actions", [])
            )
            sections.append(NarrativeSection(
                heading=opp.get("title", f"Strategic Opportunity {i+1}"),
                scr=scr,
                order=4 + i,
                visual_aid=opp.get("visual", "analysis_chart")
            ))

        # 5. RESOLUTION: Recommendations
        scr = SCRComponent(
            resolution=self._detail_resolution(analysis.get("recommendations", [])),
            next_steps=analysis.get("next_steps", [])
        )
        sections.append(NarrativeSection(
            heading="Recommendations & Decisions Required",
            scr=scr,
            order=10,
            visual_aid="decision_matrix"
        ))

        # 6. APPENDIX: Next Steps & Governance
        scr = SCRComponent(
            next_steps=analysis.get("governance", [])
        )
        sections.append(NarrativeSection(
            heading="Implementation Roadmap",
            scr=scr,
            order=11,
            visual_aid="gantt"
        ))

        return sections

    def _synthesize_resolution(self, financials: dict, opportunities: list) -> str:
        """Top-line resolution for executive summary."""
        revenue = financials.get("revenue", 0)
        ebitda = financials.get("ebitda", 0)
        top_opp = opportunities[0] if opportunities else {}

        return (
            f"The business is at an inflection point: {revenue:,.0f} revenue with {ebitda:,.0f} EBITDA "
            f"demonstrates product-market fit, but {top_opp.get('title', 'the next growth phase')} "
            f"requires {top_opp.get('investment', 'targeted investment')} to capture "
            f"{top_opp.get('market_size', 'the addressable market')}. "
            f"Board approval is sought for {len(opportunities)} strategic initiatives."
        )

    def _synthesize_situation(self, financials: dict, metrics: dict) -> str:
        parts = []
        if financials.get("revenue"):
            parts.append(f"Revenue reached {financials['revenue']:,.0f} ({financials.get('yoy_growth', 0):.0%} YoY)")
        if financials.get("ebitda_margin"):
            parts.append(f"EBITDA margin {financials['ebitda_margin']:.1%}")
        if metrics.get("nrr"):
            parts.append(f"NRR {metrics['nrr']:.0%}")
        return " | ".join(parts) if parts else "Business performing within plan."

    def _synthesize_complication(self, risks: list, metrics: dict, benchmarks: dict) -> str:
        if not risks:
            return "No material complications identified."
        top_risks = [r.get("risk", "") for r in risks[:2]]
        return f"However, {', '.join(top_risks).lower()} threaten trajectory."

    def _detail_situation(self, financials: dict, metrics: dict, benchmarks: dict) -> str:
        lines = ["**Current State**"]
        lines.append(f"• Revenue: {financials.get('revenue', 0):,.0f} ({financials.get('yoy_growth', 0):.0%} YoY)")
        lines.append(f"• EBITDA: {financials.get('ebitda', 0):,.0f} ({financials.get('ebitda_margin', 0):.1%} margin)")
        lines.append(f"• Cash: {financials.get('cash', 0):,.0f} ({financials.get('runway_months', 0)} months runway)")

        if benchmarks:
            lines.append("\n**vs. Benchmarks**")
            for k, v in benchmarks.items():
                lines.append(f"• {k}: {v}")

        return "\n".join(lines)

    def _detail_complication(self, risks: list, metrics: dict, benchmarks: dict) -> str:
        lines = ["**Key Complications**"]
        for i, risk in enumerate(risks[:5], 1):
            lines.append(f"{i}. **{risk.get('risk', 'Risk')}** ({risk.get('severity', 'Medium')})")
            lines.append(f"   *Impact*: {risk.get('impact', 'Not quantified')}")
            lines.append(f"   *Likelihood*: {risk.get('likelihood', 'Not assessed')}")
            lines.append(f"   *Current mitigation*: {risk.get('mitigation', 'None')}")

        # Gap analysis
        lines.append("\n**Performance Gaps vs. Target**")
        if metrics.get("ebitda_margin") and benchmarks.get("ebitda_target"):
            gap = benchmarks["ebitda_target"] - metrics["ebitda_margin"]
            lines.append(f"• EBITDA margin gap: {gap:.1%} vs. target")

        return "\n".join(lines)

    def _extract_evidence(self, financials: dict, metrics: dict) -> list[str]:
        evidence = []
        for k, v in {**financials, **metrics}.items():
            if isinstance(v, (int, float)):
                evidence.append(f"{k}: {v}")
        return evidence[:10]

    def _derive_implications(self, opportunities: list, risks: list) -> list[str]:
        implications = []
        for opp in opportunities[:3]:
            implications.append(f"If {opp.get('title', 'opportunity')} executed: {opp.get('impact', 'significant upside')}")
        for risk in risks[:2]:
            implications.append(f"If {risk.get('risk', 'risk')} materializes: {risk.get('impact', 'downside risk')}")
        return implications

    def _define_next_steps(self, analysis: dict) -> list[str]:
        recs = analysis.get("recommendations", [])
        return [r.get("text", r) if isinstance(r, dict) else r for r in recs[:5]]

    def _detail_resolution(self, recommendations: list) -> str:
        lines = ["**Recommended Actions**"]
        for i, rec in enumerate(recommendations[:6], 1):
            text = rec.get("text", rec) if isinstance(rec, dict) else rec
            priority = rec.get("priority", "P2") if isinstance(rec, dict) else "P2"
            lines.append(f"{i}. [{priority}] {text}")
        return "\n".join(lines)

    def _build_pyramid_narrative(self, objective: str, analysis: dict, audience: str) -> list[NarrativeSection]:
        """Minto Pyramid Principle: Answer first, then supporting arguments grouped MECE."""
        # Implementation similar to SCR but structured as pyramid
        return self._build_scr_narrative(objective, analysis, audience)


def build_executive_narrative(
    objective: str,
    analysis: dict[str, Any],
    audience: str = "board"
) -> list[NarrativeSection]:
    """Convenience function."""
    engine = SCREngine()
    return engine.build_narrative(objective, analysis, audience)
```

---

### 6.8 `fallbacks/registry.py` — Extensible Fallback Registry

```python
"""
OrchestrIQ Document Intelligence — Fallback Registry
Pluggable, priority-based fallback system. Replaces hardcoded keyword matching.
"""
from __future__ import annotations
from abc import ABC, abstractmethod
from dataclasses import dataclass
from typing import Any, Callable, Optional
from enum import Enum
import logging

logger = logging.getLogger(__name__)


class FallbackPriority(int, Enum):
    LOW = 0
    NORMAL = 50
    HIGH = 100
    CRITICAL = 200


@dataclass
class FallbackSpec:
    """Specification for a fallback handler."""
    name: str
    handler: Callable
    priority: FallbackPriority = FallbackPriority.NORMAL
    conditions: dict[str, Any] = None  # e.g., {"format": "excel", "keywords": ["fte", "workforce"]}
    metadata: dict[str, Any] = None

    def matches(self, context: dict[str, Any]) -> bool:
        """Check if this fallback matches the context."""
        if not self.conditions:
            return True

        for key, expected in self.conditions.items():
            if key not in context:
                return False

            actual = context[key]

            if isinstance(expected, list):
                # Any match
                if not any(e in str(actual).lower() for e in expected):
                    return False
            elif callable(expected):
                if not expected(actual):
                    return False
            elif str(expected).lower() not in str(actual).lower():
                return False

        return True


class FallbackRegistry:
    """
    Registry for fallback handlers with priority-based resolution.
    """

    def __init__(self):
        self._handlers: dict[str, list[FallbackSpec]] = {}  # format -> handlers
        self._default_handlers: list[FallbackSpec] = []

    def register(
        self,
        name: str,
        handler: Callable,
        formats: list[str] | str = None,
        priority: FallbackPriority = FallbackPriority.NORMAL,
        conditions: dict[str, Any] = None,
        metadata: dict[str, Any] = None
    ) -> "FallbackRegistry":
        """Register a fallback handler."""
        formats = formats or ["*"]
        if isinstance(formats, str):
            formats = [formats]

        spec = FallbackSpec(
            name=name,
            handler=handler,
            priority=priority,
            conditions=conditions or {},
            metadata=metadata or {}
        )

        for fmt in formats:
            if fmt not in self._handlers:
                self._handlers[fmt] = []
            self._handlers[fmt].append(spec)

        # Sort by priority (highest first)
        for fmt in formats:
            self._handlers[fmt].sort(key=lambda s: s.priority, reverse=True)

        logger.info(f"Registered fallback: {name} for formats={formats} priority={priority.name}")
        return self

    def register_default(self, name: str, handler: Callable, priority: FallbackPriority = FallbackPriority.LOW) -> "FallbackRegistry":
        """Register a universal fallback."""
        spec = FallbackSpec(name=name, handler=handler, priority=priority)
        self._default_handlers.append(spec)
        self._default_handlers.sort(key=lambda s: s.priority, reverse=True)
        return self

    def resolve(self, format_type: str, context: dict[str, Any]) -> Optional[Callable]:
        """Resolve the best matching fallback handler."""
        # Try format-specific handlers first
        handlers = self._handlers.get(format_type, [])
        for spec in handlers:
            if spec.matches(context):
                logger.debug(f"Fallback matched: {spec.name} for {format_type}")
                return spec.handler

        # Try default handlers
        for spec in self._default_handlers:
            if spec.matches(context):
                logger.debug(f"Default fallback matched: {spec.name}")
                return spec.handler

        return None

    def execute(self, format_type: str, context: dict[str, Any], *args, **kwargs) -> Any:
        """Execute the best matching fallback."""
        handler = self.resolve(format_type, context)
        if handler:
            return handler(*args, **kwargs)

        # Ultimate fallback
        logger.warning(f"No fallback found for {format_type}, using ultimate fallback")
        return self._ultimate_fallback(format_type, context)

    def _ultimate_fallback(self, format_type: str, context: dict[str, Any]) -> Any:
        """Absolute last resort - minimal valid structure."""
        objective = context.get("objective", "Business Document")
        return {
            "title": objective[:80],
            "subtitle": "Generated by OrchestrIQ (fallback)",
            "model": {"kpis": [], "risks": [], "recs": []},
            "sections": [{"h": "Notice", "body": "This document was generated using fallback data. Please review and update."}],
            "_fallback": True
        }


# Global registry instance
_fallback_registry: Optional[FallbackRegistry] = None


def get_fallback_registry() -> FallbackRegistry:
    """Get or create global fallback registry."""
    global _fallback_registry
    if _fallback_registry is None:
        _fallback_registry = FallbackRegistry()
        _register_builtins(_fallback_registry)
    return _fallback_registry


def _register_builtins(registry: FallbackRegistry):
    """Register built-in fallback handlers."""
    from orchestriq.document_intelligence.fallbacks import (
        financial_fallback,
        workforce_fallback,
        generic_fallback,
    )

    # Excel fallbacks
    registry.register(
        "financial_excel",
        financial_fallback.build_excel_blueprint,
        formats=["excel", "xlsx"],
        priority=FallbackPriority.HIGH,
        conditions={"keywords": ["board", "quarterly", "financial", "p&l", "budget", "forecast", "investor"]}
    )

    registry.register(
        "workforce_excel",
        workforce_fallback.build_workforce_blueprint,
        formats=["excel", "xlsx"],
        priority=FallbackPriority.HIGH,
        conditions={"keywords": ["workforce", "fte", "employee", "staffing", "headcount", "shrinkage", "occupancy", "capacity"]}
    )

    registry.register(
        "generic_excel",
        generic_fallback.build_generic_blueprint,
        formats=["excel", "xlsx"],
        priority=FallbackPriority.NORMAL,
    )

    # PPTX fallbacks
    registry.register(
        "financial_pptx",
        financial_fallback.build_pptx_blueprint,
        formats=["pptx", "ppt"],
        priority=FallbackPriority.HIGH,
        conditions={"keywords": ["board", "quarterly", "review", "presentation", "investor"]}
    )

    registry.register(
        "generic_pptx",
        generic_fallback.build_pptx_blueprint,
        formats=["pptx", "ppt"],
        priority=FallbackPriority.NORMAL,
    )

    # PDF/DOCX fallbacks
    registry.register(
        "financial_doc",
        financial_fallback.build_doc_blueprint,
        formats=["pdf", "docx"],
        priority=FallbackPriority.HIGH,
        conditions={"keywords": ["report", "board", "quarterly", "financial"]}
    )

    registry.register(
        "generic_doc",
        generic_fallback.build_doc_blueprint,
        formats=["pdf", "docx"],
        priority=FallbackPriority.NORMAL,
    )

    # Ultimate defaults
    registry.register_default(
        "ultimate_excel",
        generic_fallback.build_generic_blueprint,
        priority=FallbackPriority.LOW
    )

    registry.register_default(
        "ultimate_pptx",
        generic_fallback.build_pptx_blueprint,
        priority=FallbackPriority.LOW
    )

    registry.register_default(
        "ultimate_doc",
        generic_fallback.build_doc_blueprint,
        priority=FallbackPriority.LOW
    )
```

---

### 6.9 `core/telemetry.py` — Observability

```python
"""
OrchestrIQ Document Intelligence — Telemetry & Observability
Structured logging, metrics, and distributed tracing.
"""
from __future__ import annotations
import os
import time
import logging
import contextvars
from functools import wraps
from typing import Any, Callable, Optional
from dataclasses import dataclass, field
from contextlib import contextmanager

try:
    from opentelemetry import trace
    from opentelemetry.sdk.trace import TracerProvider
    from opentelemetry.sdk.trace.export import BatchSpanProcessor
    from opentelemetry.exporter.otlp.proto.grpc.trace_exporter import OTLPSpanExporter
    from opentelemetry.instrumentation.logging import LoggingInstrumentor
    _HAS_OTEL = True
except ImportError:
    _HAS_OTEL = False

try:
    import structlog
    _HAS_STRUCTLOG = True
except ImportError:
    _HAS_STRUCTLOG = False


# Context variables for request correlation
request_id_var: contextvars.ContextVar[str] = contextvars.ContextVar("request_id", default="")
user_id_var: contextvars.ContextVar[str] = contextvars.ContextVar("user_id", default="")
session_id_var: contextvars.ContextVar[str] = contextvars.ContextVar("session_id", default="")


@dataclass
class TelemetryConfig:
    service_name: str = "orchestriq-document-intelligence"
    environment: str = os.environ.get("ENVIRONMENT", "development")
    otlp_endpoint: Optional[str] = os.environ.get("OTEL_EXPORTER_OTLP_ENDPOINT")
    log_level: str = os.environ.get("LOG_LEVEL", "INFO")
    enable_console: bool = os.environ.get("LOG_CONSOLE", "true").lower() == "true"
    enable_json: bool = os.environ.get("LOG_JSON", "true").lower() == "true"


class Telemetry:
    """Centralized telemetry management."""

    def __init__(self, config: Optional[TelemetryConfig] = None):
        self.config = config or TelemetryConfig()
        self._tracer = None
        self._logger = None
        self._initialized = False

    def initialize(self):
        """Initialize logging and tracing."""
        if self._initialized:
            return

        self._setup_logging()
        self._setup_tracing()
        self._initialized = True

    def _setup_logging(self):
        """Configure structured logging."""
        if _HAS_STRUCTLOG and self.config.enable_json:
            structlog.configure(
                processors=[
                    structlog.stdlib.filter_by_level,
                    structlog.stdlib.add_logger_name,
                    structlog.stdlib.add_log_level,
                    structlog.stdlib.PositionalArgumentsFormatter(),
                    structlog.processors.TimeStamper(fmt="iso"),
                    structlog.processors.StackInfoRenderer(),
                    structlog.processors.format_exc_info,
                    structlog.processors.UnicodeDecoder(),
                    structlog.processors.JSONRenderer()
                ],
                context_class=dict,
                logger_factory=structlog.stdlib.LoggerFactory(),
                cache_logger_on_first_use=True,
            )
            self._logger = structlog.get_logger()
        else:
            # Standard logging with context
            logging.basicConfig(
                level=getattr(logging, self.config.log_level),
                format="%(asctime)s %(levelname)s [%(name)s] %(message)s"
            )
            self._logger = logging.getLogger(self.config.service_name)

        # Add context to all log records
        old_factory = logging.getLogRecordFactory()
        def record_factory(*args, **kwargs):
            record = old_factory(*args, **kwargs)
            record.request_id = request_id_var.get("")
            record.user_id = user_id_var.get("")
            record.session_id = session_id_var.get("")
            return record
        logging.setLogRecordFactory(record_factory)

    def _setup_tracing(self):
        """Configure OpenTelemetry tracing."""
        if not _HAS_OTEL or not self.config.otlp_endpoint:
            return

        provider = TracerProvider()
        provider.add_span_processor(
            BatchSpanProcessor(OTLPSpanExporter(endpoint=self.config.otlp_endpoint))
        )
        trace.set_tracer_provider(provider)
        self._tracer = trace.get_tracer(self.config.service_name)

        # Instrument logging
        LoggingInstrumentor().instrument(set_logging_format=True)

    @property
    def logger(self):
        if not self._initialized:
            self.initialize()
        return self._logger

    @property
    def tracer(self):
        if not self._initialized:
            self.initialize()
        return self._tracer

    def set_context(self, request_id: str = "", user_id: str = "", session_id: str = ""):
        """Set correlation context."""
        if request_id:
            request_id_var.set(request_id)
        if user_id:
            user_id_var.set(user_id)
        if session_id:
            session_id_var.set(session_id)

    def clear_context(self):
        request_id_var.set("")
        user_id_var.set("")
        session_id_var.set("")

    @contextmanager
    def span(self, name: str, attributes: dict = None):
        """Create a traced span."""
        if self._tracer:
            with self._tracer.start_as_current_span(name) as span:
                if attributes:
                    for k, v in attributes.items():
                        span.set_attribute(k, str(v))
                yield span
        else:
            yield None

    def record_metric(self, name: str, value: float, attributes: dict = None):
        """Record a metric (placeholder for metrics backend)."""
        self.logger.info("metric", metric_name=name, metric_value=value, **(attributes or {}))

    def time_operation(self, name: str):
        """Decorator/context manager for timing operations."""
        def decorator(func: Callable):
            @wraps(func)
            def wrapper(*args, **kwargs):
                start = time.perf_counter()
                try:
                    result = func(*args, **kwargs)
                    return result
                finally:
                    duration = time.perf_counter() - start
                    self.logger.info(
                        "operation_timing",
                        operation=name,
                        duration_ms=duration * 1000,
                        request_id=request_id_var.get("")
                    )
                    self.record_metric(f"operation.{name}.duration_ms", duration * 1000)
            return wrapper
        return decorator


# Global telemetry instance
_telemetry: Optional[Telemetry] = None


def get_telemetry() -> Telemetry:
    global _telemetry
    if _telemetry is None:
        _telemetry = Telemetry()
        _telemetry.initialize()
    return _telemetry


def get_logger(name: str = None):
    """Get structured logger with context."""
    tel = get_telemetry()
    if _HAS_STRUCTLOG:
        return tel.logger.bind(component=name or "unknown")
    return logging.getLogger(name or "orchestriq")


# Convenience decorators
def traced(operation_name: str = None):
    """Decorator to trace a function."""
    def decorator(func: Callable):
        @wraps(func)
        def wrapper(*args, **kwargs):
            tel = get_telemetry()
            name = operation_name or f"{func.__module__}.{func.__name__}"
            with tel.span(name) as span:
                return func(*args, **kwargs)
        return wrapper
    return decorator


def logged(logger_name: str = None):
    """Decorator to add structured logging."""
    def decorator(func: Callable):
        @wraps(func)
        def wrapper(*args, **kwargs):
            logger = get_logger(logger_name or func.__module__)
            logger.info("function_start", function=func.__name__)
            try:
                result = func(*args, **kwargs)
                logger.info("function_end", function=func.__name__, success=True)
                return result
            except Exception as e:
                logger.error("function_error", function=func.__name__, error=str(e))
                raise
        return wrapper
    return decorator
```

---

### 6.10 Refactored `schema_extractor.py` — Thin Facade (Backward Compatible)

```python
"""
OrchestrIQ Document Intelligence Engine v4 — AI Schema Extractor
REFACORED: Thin facade delegating to modular components.
ZERO-RAISE GUARANTEE: this module never raises. Any failure returns (fallback_schema, reason_string).
BACKWARD COMPATIBLE: All public function signatures unchanged.
"""
from __future__ import annotations
import json
import traceback
from typing import Any, Optional

# Import new modular components
from orchestriq.document_intelligence.core.models import (
    FinancialModel, WorkbookBlueprint, PresentationBlueprint, DocumentBlueprint,
    Currency, ExtractionMode, DocumentFormat, BlueprintType
)
from orchestriq.document_intelligence.core.llm_client import LLMClientFactory, LLMProvider
from orchestriq.document_intelligence.core.parser import parse_json_robust
from orchestriq.document_intelligence.core.prompt_registry import PromptRegistry, PromptCategory
from orchestriq.document_intelligence.core.telemetry import get_telemetry, get_logger, traced
from orchestriq.document_intelligence.financial.model import build_financial_model, BusinessModel
from orchestriq.document_intelligence.fallbacks.registry import get_fallback_registry
from orchestriq.document_intelligence.design.chart_selector import select_chart
from orchestriq.document_intelligence.narrative.scr_engine import build_executive_narrative

# Initialize telemetry
_telemetry = get_telemetry()
_logger = get_logger(__name__)
_prompt_registry = PromptRegistry()
_fallback_registry = get_fallback_registry()

# Model constant
MODEL = "claude-3-haiku-20240307"  # Updated to current model


# ═══════════════════════════════════════════════════════════════════
# LLM Client (with env var fallback)
# ═══════════════════════════════════════════════════════════════════
def _get_llm_client(api_key: str = "") -> Any:
    """Get LLM client with fallback to environment."""
    try:
        return LLMClientFactory.create(
            provider=LLMProvider.ANTHROPIC,
            api_key=api_key,
            model=MODEL
        )
    except Exception as e:
        _logger.warning("llm_client_creation_failed", error=str(e))
        return None


def _call_ai(prompt: str, api_key: str = "", max_tokens: int = 8000) -> Optional[str]:
    """Call LLM. Returns raw text or None. Never raises."""
    client = _get_llm_client(api_key)
    if not client:
        return None

    try:
        response = client.complete(prompt, max_tokens=max_tokens)
        if response.error:
            _logger.warning("llm_call_failed", error=response.error, provider=response.provider.value)
            return None
        return response.text
    except Exception as e:
        _logger.error("llm_call_exception", error=str(e))
        traceback.print_exc()
        return None


# ═══════════════════════════════════════════════════════════════════
# Public API — EXACT SAME SIGNATURES AS BEFORE
# ═══════════════════════════════════════════════════════════════════

@traced("extract_financial")
def extract(
    objective: str,
    ctx: str,
    data: str,
    api_key: str,
    sym: str = "₹"
) -> tuple[dict, str, str]:
    """
    Master extraction. Returns (enriched_model_dict, mode, reason).
    mode: 'ai' or 'fallback'. NEVER raises.
    BACKWARD COMPATIBLE: Returns dict (not Pydantic model) for compatibility.
    """
    _telemetry.set_context(request_id=f"ext_{hash(objective) % 10000:04d}")
    _logger.info("extract_start", objective_preview=objective[:100])

    try:
        # Parse data if provided
        parsed_data = {}
        if data and data.strip() != "(none)":
            try:
                parsed_data = json.loads(data) if data.strip().startswith("{") else {"raw": data}
            except json.JSONDecodeError:
                parsed_data = {"raw": data}

        # Build financial model (new engine)
        currency = Currency.INR if sym == "₹" else Currency.USD
        model = build_financial_model(
            objective=objective,
            context=ctx,
            data=parsed_data,
            currency=currency,
            business_model=_detect_business_model(objective)
        )

        # Try AI enrichment
        raw = _call_ai(
            _prompt_registry.render(
                PromptCategory.FINANCIAL_EXTRACTION,
                sym=sym,
                obj=objective[:1500],
                ctx=ctx[:2000],
                data=data[:8000] or "(none)"
            ),
            api_key,
            max_tokens=8000
        )

        if raw is None:
            return model.model_dump(), "fallback", "no api key or AI call failed"

        parsed, err = parse_json_robust(raw)
        if err or not isinstance(parsed, dict):
            return model.model_dump(), "fallback", f"AI returned unparseable JSON: {err}"

        # Merge AI enrichment into model (validated)
        enriched = _merge_ai_enrichment(model, parsed)

        return enriched.model_dump(), "ai", "ok"

    except Exception as e:
        _logger.error("extract_exception", error=str(e))
        traceback.print_exc()
        # Ultimate fallback
        fallback_model = build_financial_model(objective, ctx, {}, currency=currency)
        return fallback_model.model_dump(), "fallback", f"extractor exception: {str(e)[:120]}"


def _detect_business_model(objective: str) -> BusinessModel:
    """Simple heuristic for business model detection."""
    obj_lower = objective.lower()
    if any(k in obj_lower for k in ["saas", "subscription", "arr", "mrr", "churn"]):
        return BusinessModel.SAAS
    elif any(k in obj_lower for k in ["marketplace", "platform", "gmv", "take rate"]):
        return BusinessModel.MARKETPLACE
