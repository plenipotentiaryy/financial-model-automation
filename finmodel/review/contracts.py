"""Versioned, JSON-serializable review contracts. Decimal values travel as strings."""
from dataclasses import dataclass, field, asdict
from hashlib import sha256
import json

SCHEMA_VERSION = 1
POLICY_VERSION = 'financial-evidence-v1'
MODEL = 'jev-1.13.0'
CHECKS = ('metric', 'number_binding', 'period', 'units', 'entity_scope', 'context')


def digest(value):
    return sha256(json.dumps(value, sort_keys=True, ensure_ascii=False, separators=(',', ':')).encode()).hexdigest()


@dataclass
class Candidate:
    source_file: str
    source_sha256: str
    locator: dict
    excerpt: str
    raw_label: str
    raw_value: str
    metric: str | None = None
    company_id: str = ''
    company_name: str = ''
    fiscal_year: int | None = None
    period_type: str = ''
    currency: str = ''
    monetary_unit: str = ''
    scope: str = ''
    headers: list = field(default_factory=list)
    normalized_value: str | None = None
    value_status: str = 'missing'
    extraction_flags: list = field(default_factory=list)
    recipe: dict | None = None
    components: list = field(default_factory=list)
    schema_version: int = SCHEMA_VERSION
    revision: int = 1
    candidate_id: str = ''

    def __post_init__(self):
        if not self.candidate_id:
            self.candidate_id = digest([self.source_file, self.source_sha256, self.locator, self.raw_label, self.raw_value])[:24]

    def payload(self):
        return asdict(self)


@dataclass
class VerificationResult:
    candidate_id: str
    revision: int
    status: str
    reason: str
    probabilities: dict = field(default_factory=dict)
    requested_model: str = MODEL
    actual_model: str | None = None
    response: dict = field(default_factory=dict)
    usage: dict = field(default_factory=dict)
    cost_usd: str = '0'
    cached: bool = False
    mode: str = 'live'
    policy_version: str = POLICY_VERSION
    schema_version: int = SCHEMA_VERSION


@dataclass
class ReviewDecision:
    candidate_id: str
    revision: int
    action: str
    reason: str
    reviewer: str = 'local reviewer'
    schema_version: int = SCHEMA_VERSION
