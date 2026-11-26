from dataclasses import dataclass, field
from datetime import datetime
from typing import Set, Dict, Any
import pandas as pd

OPEN_OMS_STATUSES: Set[str] = {
    "ASSIGNED",
    "IN PROGRESS",
    "ON HOLD",
    "PENDING APPROVAL",
    "PENDING CET APPROVAL",
    "OPEN",
    "RE-OPENED",
}

CLOSED_OMS_STATUSES: Set[str] = {
    "CLOSED",
    "RESOLVED",
    "CANCELLED",
    "REJECTED",
}

DOWN_FIR_STATUSES: Set[str] = {
    "FAULTED",
    "UNAVAILABLE",
    "OFFLINE",
    "DOWN",
}

CATEGORY_GAP: str = "Ticket Generation Gap"
CATEGORY_PHANTOM: str = "Phantom Resolution"
CATEGORY_ACTIVE: str = "Active Ticket Verified"
CATEGORY_HEALTHY: str = "Healthy / Cleared"
CATEGORY_UNTRACKED: str = "OMS Untracked Open Tickets"

RISK_CRITICAL: str = "Critical"
RISK_HIGH: str = "High"
RISK_MEDIUM: str = "Medium"
RISK_LOW: str = "Low"


@dataclass
class ReconciliationSummary:
    total_fir: int
    in_maintenance: int
    available_cleared: int
    ticket_gaps: int
    phantom_resolutions: int
    active_verified: int
    coverage_gap_rate: float
    untracked_oms_open: int
    timestamp: str = field(default_factory=lambda: datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    execution_time_ms: float = 0.0

    def to_dict(self) -> Dict[str, Any]:
        return {
            "total_fir": self.total_fir,
            "in_maintenance": self.in_maintenance,
            "available_cleared": self.available_cleared,
            "ticket_gaps": self.ticket_gaps,
            "phantom_resolutions": self.phantom_resolutions,
            "active_verified": self.active_verified,
            "coverage_gap_rate": self.coverage_gap_rate,
            "untracked_oms_open": self.untracked_oms_open,
            "timestamp": self.timestamp,
            "execution_time_ms": self.execution_time_ms,
        }


@dataclass
class ReconciliationResult:
    summary: ReconciliationSummary
    reconciled_df: pd.DataFrame
    gaps_df: pd.DataFrame
    phantoms_df: pd.DataFrame
    active_df: pd.DataFrame
    healthy_df: pd.DataFrame
    untracked_oms_df: pd.DataFrame
