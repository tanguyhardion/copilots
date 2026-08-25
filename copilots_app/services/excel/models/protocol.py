"""Action Protocol models defining LLM action structures and execution results."""

from enum import Enum
from typing import List, Optional, Any, Dict
from pydantic import BaseModel, Field


class ActionIntent(str, Enum):
    MODIFY_WORKBOOK = "modify_workbook"
    QUERY_WORKBOOK = "query_workbook"
    GENERATE_WORKBOOK = "generate_workbook"
    NO_ACTION = "no_action"


class ExecutionStatus(str, Enum):
    SUCCESS = "success"
    PARTIAL_SUCCESS = "partial_success"
    VALIDATION_FAILED = "validation_failed"
    EXECUTION_FAILED = "execution_failed"


class ActionItem(BaseModel):
    """Represents an individual action to execute on Excel."""

    action: str = Field(..., description="Action identifier (e.g. add_column, create_chart)")
    sheet: Optional[str] = Field(None, description="Target worksheet name")
    table: Optional[str] = Field(None, description="Target table name")
    column: Optional[str] = Field(None, description="Target column name")
    new_name: Optional[str] = Field(None, description="New name for sheet or column")
    formula: Optional[str] = Field(None, description="Semantic formula (e.g. {Profit}/{Revenue})")
    values: Optional[List[Any]] = Field(None, description="Data values list")
    search_query: Optional[str] = Field(None, description="Query string for search actions")
    chart: Optional[str] = Field(None, description="Chart identifier or title")
    chart_type: Optional[str] = Field(None, description="Type of chart (bar, line, pie, column)")
    data_range: Optional[str] = Field(None, description="Cell range or table name for chart data")
    params: Dict[str, Any] = Field(default_factory=dict, description="Additional custom action parameters")


class ActionProtocol(BaseModel):
    """Parsed excel-action block payload from LLM response."""

    intent: ActionIntent = Field(ActionIntent.NO_ACTION, description="Primary intent")
    actions: List[ActionItem] = Field(default_factory=list, description="Ordered action items")
    explanation: Optional[str] = Field(None, description="Optional explanation or intent summary")


class ValidationResult(BaseModel):
    """Result of pre-execution validation engine."""

    is_valid: bool = Field(True, description="True if valid to execute")
    errors: List[str] = Field(default_factory=list, description="Blocking validation errors")
    warnings: List[str] = Field(default_factory=list, description="Non-blocking warnings")
    action_previews: List[str] = Field(default_factory=list, description="Human-readable execution preview descriptions")


class ExecutionResult(BaseModel):
    """Structured report returned after executing actions."""

    status: ExecutionStatus = Field(ExecutionStatus.SUCCESS, description="Final execution outcome")
    actions_executed: int = Field(0, description="Count of actions executed")
    objects_modified: List[str] = Field(default_factory=list, description="Human-readable names of objects updated")
    warnings: List[str] = Field(default_factory=list, description="Warnings generated during execution")
    errors: List[str] = Field(default_factory=list, description="Error messages if execution failed")
    details: List[Dict[str, Any]] = Field(default_factory=list, description="Detailed per-action log items")
    query_results: Optional[List[Dict[str, Any]]] = Field(None, description="Results if intent was query_workbook")
