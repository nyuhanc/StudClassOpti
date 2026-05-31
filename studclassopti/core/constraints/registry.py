"""The catalog of constraints the model and GUI walk.

``REGISTRY`` is ordered to match the original numbering (5 and 7 absent; 5 is
re-added here, off by default). ``build_model`` iterates it; the GUI renders it.
"""

from __future__ import annotations

import pandas as pd

from ..config import SolverConfig
from ..data import DataError
from .base import Constraint
from .c01_class_size import ClassSize
from .c02_lang_capacity import LangCapacity
from .c03_natsci_capacity import NatSciCapacity
from .c04_distinct_sciences import DistinctSciences
from .c05_schoolmate import Schoolmate
from .c06_best_pair_class import BestPairClass
from .c08_male_split import MaleSplit
from .c09_force_top_language import ForceTopLanguage
from .c10_subject_capacity import SubjectCapacity
from .c11_require_subject import RequireSubjectForTopPriority
from .c12_exclude_subject import ExcludeSubjectForLowPriority
from .c13_subject_pairing import SubjectPairing
from .c14_language_exclusion import LanguageExclusion

REGISTRY: list[Constraint] = [
    ClassSize(),
    LangCapacity(),
    NatSciCapacity(),
    DistinctSciences(),
    Schoolmate(),
    BestPairClass(),
    MaleSplit(),
    ForceTopLanguage(),
    SubjectCapacity(),
    RequireSubjectForTopPriority(),
    ExcludeSubjectForLowPriority(),
    SubjectPairing(),
    LanguageExclusion(),
]


def enabled_constraints(config: SolverConfig) -> list[Constraint]:
    """The constraints whose toggle is on in ``config``."""
    return [c for c in REGISTRY if getattr(config.constraints, c.enabled_field)]


def validate_constraints(data: pd.DataFrame, config: SolverConfig) -> list[DataError]:
    """Aggregate the ``validate`` output of every enabled constraint."""
    errors: list[DataError] = []
    for c in enabled_constraints(config):
        errors.extend(c.validate(data, config))
    return errors
