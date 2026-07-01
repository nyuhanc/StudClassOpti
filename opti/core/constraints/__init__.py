"""Self-describing constraint plugins and their registry."""

from .base import Constraint, ModelContext, Parameter
from .registry import REGISTRY, enabled_constraints, validate_constraints
