"""
MS Project XML → PROGRAMME list parser
========================================
Extracts activity-level tasks (OutlineLevel 5) from an MS Project .xml export.
Returns a list of dicts matching the PROGRAMME format expected by lookahead_engine.

For each task it derives:
  - name, prog_start, prog_finish from the XML
  - target_wd by counting actual working days in the WD calendar
  - section / zone inferred from WBS/outline structure (or left for user to assign)
  - subtasks: child tasks at OutlineLevel 6
"""

import xml.etree.ElementTree as ET
from datetime import datetime, date
from typing import Optional


NS = "http://schemas.microsoft.com/project"


def _tag(name):
    return f"{{{NS}}}{name}"


def parse_msp_xml(xml_path: str, working: set, outline_level: int = 5) -> list[dict]:
    """
    Parse MS Project XML and return a PROGRAMME-compatible list.

    Args:
        xml_path:      Path to the .xml file exported from MS Project
        working:       Set of working date objects (from load_calendar)
        outline_level: OutlineLevel of activity rows (default 5)

    Returns:
        List of activity dicts with keys:
            section, zone, name, target_wd, prog_start, prog_finish, subtasks
    """
    # Build UID → task mapping via streaming parse
    tasks_by_uid = {}
    uid_order = []
    current = {}

    for event, elem in ET.iterparse(xml_path, events=("start", "end")):
        if event == "start" and elem.tag == _tag("Task"):
            current = {}
        elif event == "end":
            tag = elem.tag.replace(f"{{{NS}}}", "")
            if tag in ("UID", "ID", "Name", "OutlineLevel", "Summary",
                       "Start", "Finish", "OutlineNumber"):
                current[tag] = elem.text
            elif tag == "Task":
                uid = current.get("UID")
                if uid:
                    tasks_by_uid[uid] = dict(current)
                    uid_order.append(uid)
                current = {}

    # Build ordered list of tasks by their sequential ID
    all_tasks = [tasks_by_uid[uid] for uid in uid_order if uid in tasks_by_uid]
    all_tasks.sort(key=lambda t: int(t.get("ID", 0)))

    # Index tasks by outline number for hierarchy lookup
    by_outline = {}
    for t in all_tasks:
        on = t.get("OutlineNumber", "")
        if on:
            by_outline[on] = t

    def _parent_outline(outline_num: str) -> Optional[str]:
        parts = outline_num.rsplit(".", 1)
        return parts[0] if len(parts) > 1 else None

    def _ol(t) -> int:
        try:
            return int(t.get("OutlineLevel", 99))
        except (ValueError, TypeError):
            return 99

    def _parse_date(s: str) -> Optional[date]:
        if not s:
            return None
        for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%d"):
            try:
                return datetime.strptime(s[:19], fmt).date()
            except ValueError:
                continue
        return None

    def _count_wd(start: date, finish: date) -> int:
        if not start or not finish or finish < start:
            return 0
        count, cur = 0, start
        from datetime import timedelta
        while cur <= finish:
            if cur in working:
                count += 1
            cur += timedelta(days=1)
        return count

    def _infer_section_zone(task: dict, all_tasks: list) -> tuple[str, str]:
        """Walk up the outline hierarchy to infer section and zone labels."""
        on = task.get("OutlineNumber", "")
        section, zone = "GENERAL", ""

        # Try grandparent (OL 3) as section, parent (OL 4) as zone
        parts = on.split(".")
        if len(parts) >= 4:
            grandparent_on = ".".join(parts[:3])
            parent_on = ".".join(parts[:4])
            gp = by_outline.get(grandparent_on)
            p  = by_outline.get(parent_on)
            if gp:
                section = gp.get("Name", "GENERAL").strip().upper()
            if p:
                zone = p.get("Name", "").strip()
        elif len(parts) >= 3:
            parent_on = ".".join(parts[:3])
            p = by_outline.get(parent_on)
            if p:
                section = p.get("Name", "GENERAL").strip().upper()
        return section, zone

    # Collect activities at the target outline level
    programme = []
    for t in all_tasks:
        if _ol(t) != outline_level:
            continue
        if t.get("Summary", "0") == "1":
            continue  # skip summary rows at this level

        name = t.get("Name", "").strip()
        if not name or name.startswith("Task "):
            continue

        start_date  = _parse_date(t.get("Start", ""))
        finish_date = _parse_date(t.get("Finish", ""))
        target_wd   = _count_wd(start_date, finish_date) if start_date and finish_date else 0

        # Collect sub-tasks (OL 6 children)
        my_on = t.get("OutlineNumber", "")
        subtasks = []
        for child in all_tasks:
            if _ol(child) != outline_level + 1:
                continue
            child_on = child.get("OutlineNumber", "")
            if child_on.startswith(my_on + ".") and child_on.count(".") == my_on.count(".") + 1:
                child_name = child.get("Name", "").strip()
                if child_name:
                    subtasks.append(child_name)

        # Default subtask if none found
        if not subtasks:
            subtasks = ["Preparation", "Works", "Clean Up & Inspect"]

        section, zone = _infer_section_zone(t, all_tasks)

        programme.append({
            "section":     section,
            "zone":        zone,
            "name":        name,
            "target_wd":   target_wd,
            "prog_start":  start_date.isoformat() if start_date else "",
            "prog_finish": finish_date.isoformat() if finish_date else "",
            "subtasks":    subtasks,
        })

    return programme


def summarise(programme: list[dict]) -> dict:
    """Return summary stats for display in the UI."""
    from datetime import date as dclass
    today = date.today()
    total = len(programme)
    in_progress, starting_soon, future, past = 0, 0, 0, 0
    sections = set()
    zones = set()
    for act in programme:
        sections.add(act["section"])
        zones.add(act["zone"])
        ps = datetime.strptime(act["prog_start"], "%Y-%m-%d").date() if act["prog_start"] else None
        pf = datetime.strptime(act["prog_finish"], "%Y-%m-%d").date() if act["prog_finish"] else None
        if ps and pf:
            if pf < today:
                past += 1
            elif ps <= today:
                in_progress += 1
            elif (ps - today).days <= 14:
                starting_soon += 1
            else:
                future += 1
    return {
        "total": total,
        "in_progress": in_progress,
        "starting_soon": starting_soon,
        "future": future,
        "past": past,
        "sections": sorted(sections),
        "zones": sorted(zones),
    }
