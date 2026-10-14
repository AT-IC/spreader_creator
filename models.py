from dataclasses import dataclass
from typing import List


@dataclass
class ParsedData:
    shifts: List[dict]
    leave_list: List[str]
    notes_list: List[str]
