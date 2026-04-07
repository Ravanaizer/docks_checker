from dataclasses import dataclass
from enum import Enum


class Severity(Enum):
    CRITICAL = "CRITICAL"
    WARNING = "WARNING"
    INFO = "INFO"


@dataclass
class ValidationError:
    rule: str
    message: str
    severity: Severity
    location: str = ""
