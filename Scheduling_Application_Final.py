import os
from flask import Flask, render_template_string, request, redirect, url_for, flash, session, send_from_directory, send_file
from functools import wraps
import pyodbc
app = Flask(__name__)
app.secret_key = "dev-secret-change-me"
########################################################################################################################





########################################################################################################################
########################################################################################################################
#######################################################################################################################
"""DATABASE CONNECTION"""
ACCESS_DB_PATH = r"C:\Users\Brooks\OneDrive\Desktop\Scheduling Application - Final.accdb"
ACCESS_DRIVER = r"{Microsoft Access Driver (*.mdb, *.accdb)}"
DEFAULT_NOT_STARTED_STATUS_NAME = "Not Started"
INSTALLATION_RESOURCE_TYPE_ID = 13
TECH_OPS_TASK_ID = 87
OVERHEAD_TASK_ID = 88
INDIRECT_TASK_IDS = [TECH_OPS_TASK_ID, OVERHEAD_TASK_ID]
"""FILE CONNECTION"""
TASK_FILES_ROOT = r"C:\Users\Brooks\OneDrive\Desktop\Task Files"
########################################################################################################################
########################################################################################################################
########################################################################################################################





########################################################################################################################
########################################################################################################################
########################################################################################################################
"""DATABASE HELPERS"""
########################################################################################################################
########################################################################################################################
########################################################################################################################
def get_connection():
    conn_str = (
        rf"DRIVER={ACCESS_DRIVER};"
        rf"DBQ={ACCESS_DB_PATH};"
    )
    return pyodbc.connect(conn_str)
########################################################################################################################
def fetch_all(query, params=None):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(query, params or [])
    rows = cur.fetchall()
    conn.close()
    return rows
########################################################################################################################
def fetch_one(query, params=None):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(query, params or [])
    row = cur.fetchone()
    conn.close()
    return row
########################################################################################################################
def execute_sql(query, params=None):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(query, params or [])
    conn.commit()
    conn.close()
########################################################################################################################
def get_resource_types():
    return fetch_all(
        """
        SELECT resource_type_id, resource_description
        FROM resource_types
        ORDER BY resource_description
        """
    )
########################################################################################################################
def get_resources():
    return fetch_all(
        """
        SELECT resource_id, resource_name, resource_types, default_hours_per_day, resource_members
        FROM resources
        ORDER BY resource_name
        """
    )
########################################################################################################################
def get_templates():
    return fetch_all(
        """
        SELECT template_id, template_name, template_description
        FROM tasking_templates
        ORDER BY template_name
        """
    )
########################################################################################################################
def get_resource_name_map():
    rows = fetch_all(
        """
        SELECT resource_id, resource_name
        FROM resources
        """
    )
    return {row.resource_id: row.resource_name for row in rows}
########################################################################################################################
def get_resource_type_name_map():
    rows = fetch_all(
        """
        SELECT resource_type_id, resource_description
        FROM resource_types
        """
    )
    return {row.resource_type_id: row.resource_description for row in rows}
########################################################################################################################
def get_status_name_map():
    rows = fetch_all(
        """
        SELECT status_id, status_names
        FROM task_status
        """
    )
    return {row.status_id: row.status_names for row in rows}
########################################################################################################################
def get_status_id_by_name(status_name):
    row = fetch_one(
        """
        SELECT TOP 1 status_id
        FROM task_status
        WHERE status_names = ?
        """,
        [status_name],
    )
    return row.status_id if row else None
########################################################################################################################
def is_yes_no_true(value):
    if value is True:
        return True
    if value is False or value is None:
        return False

    text_value = str(value).strip().lower()
    return text_value in ["true", "yes", "y", "1", "-1"]
########################################################################################################################
def get_logged_in_employee():
    employee_id = session.get("employee_id")

    if not employee_id:
        return None

    return fetch_one(
        """
        SELECT employee_id, employee_first_name, employee_last_name, resource_id,
               email_address, [office], [installer],[admin]
        FROM employees
        WHERE employee_id = ?
        """,
        [employee_id],
    )
########################################################################################################################
def is_office_employee(employee_row):
    if not employee_row:
        return False
    return is_yes_no_true(employee_row.office)
########################################################################################################################
def is_admin_employee(employee_row):
    if not employee_row:
        return False

    return is_yes_no_true(getattr(employee_row, "admin", False))
########################################################################################################################
def login_required(view_func):
    @wraps(view_func)
    def wrapper(*args, **kwargs):
        if "employee_id" not in session:
            return redirect(url_for("login"))
        return view_func(*args, **kwargs)
    return wrapper
########################################################################################################################
def office_required(view_func):
    @wraps(view_func)
    def wrapper(*args, **kwargs):
        employee = get_logged_in_employee()

        if not employee:
            return redirect(url_for("login"))

        if not is_office_employee(employee):
            return redirect(url_for("employee_tasks"))

        return view_func(*args, **kwargs)
    return wrapper
########################################################################################################################
def get_status_id_by_name_flexible(status_names):
    placeholders = ",".join(["?"] * len(status_names))
    row = fetch_one(
        f"""
        SELECT TOP 1 status_id
        FROM task_status
        WHERE LCase(status_names) IN ({placeholders})
        """,
        [s.lower() for s in status_names],
    )
    return row.status_id if row else None
########################################################################################################################
def get_open_punch_for_employee(employee_id):
    return fetch_one(
        """
        SELECT TOP 1 time_entry_ID, task_id, punch_in
        FROM time_punches
        WHERE employee_id = ?
          AND task_id IS NOT NULL
          AND punch_out IS NULL
        ORDER BY punch_in DESC
        """,
        [employee_id],
    )
########################################################################################################################
def get_open_day_punch_for_employee(employee_id):
    return fetch_one(
        """
        SELECT TOP 1 time_entry_ID, punch_in
        FROM time_punches
        WHERE employee_id = ?
          AND task_id IS NULL
          AND punch_out IS NULL
        ORDER BY punch_in DESC
        """,
        [employee_id],
    )
########################################################################################################################
########################################################################################################################
########################################################################################################################





########################################################################################################################
########################################################################################################################
########################################################################################################################
"""SCHEDULING HELPERS"""
########################################################################################################################
########################################################################################################################
########################################################################################################################
from datetime import datetime, date, time, timedelta
import math
########################################################################################################################
WORKDAY_START_HOUR = 7
LUNCH_START_HOUR = 12
LUNCH_END_HOUR = 13
WORKDAY_END_HOUR = 16
HOURS_PER_WORKDAY = 8.0
SPECIAL_ORDER_GAP_DAYS = 4
########################################################################################################################
def count_resource_members(resource_members_text):
    if resource_members_text is None:
        return 1
    text_value = str(resource_members_text).strip()
    if text_value == "":
        return 1
    return len([x for x in text_value.split(",") if x.strip() != ""])
########################################################################################################################
def combine_date_and_workday_end(dt_value):
    if isinstance(dt_value, datetime):
        return datetime.combine(dt_value.date(), time(WORKDAY_END_HOUR, 0))
    if isinstance(dt_value, date):
        return datetime.combine(dt_value, time(WORKDAY_END_HOUR, 0))
    return datetime.combine(datetime.now().date(), time(WORKDAY_END_HOUR, 0))
########################################################################################################################
def is_weekend(dt_value):
    return dt_value.weekday() >= 5
########################################################################################################################
def previous_workday(dt_value):
    current = dt_value - timedelta(days=1)
    while current.weekday() >= 5:
        current -= timedelta(days=1)
    return current
########################################################################################################################
def next_workday(dt_value):
    current = dt_value + timedelta(days=1)
    while current.weekday() >= 5:
        current += timedelta(days=1)
    return current
########################################################################################################################
def move_to_previous_work_time(dt_value):
    current = dt_value

    while is_weekend(current):
        current = datetime.combine(previous_workday(current.date()), time(WORKDAY_END_HOUR, 0))

    current_time = current.time()

    if current_time > time(WORKDAY_END_HOUR, 0):
        return datetime.combine(current.date(), time(WORKDAY_END_HOUR, 0))

    if current_time > time(LUNCH_START_HOUR, 0) and current_time < time(LUNCH_END_HOUR, 0):
        return datetime.combine(current.date(), time(LUNCH_START_HOUR, 0))

    if current_time == time(LUNCH_END_HOUR, 0):
        return datetime.combine(current.date(), time(LUNCH_START_HOUR, 0))

    if current_time < time(WORKDAY_START_HOUR, 0):
        prev_day = previous_workday(current.date())
        return datetime.combine(prev_day, time(WORKDAY_END_HOUR, 0))

    return current
########################################################################################################################
def move_to_next_work_time(dt_value):
    current = dt_value

    while is_weekend(current):
        current = datetime.combine(next_workday(current.date()), time(WORKDAY_START_HOUR, 0))

    current_time = current.time()

    if current_time < time(WORKDAY_START_HOUR, 0):
        return datetime.combine(current.date(), time(WORKDAY_START_HOUR, 0))

    if current_time >= time(LUNCH_START_HOUR, 0) and current_time < time(LUNCH_END_HOUR, 0):
        return datetime.combine(current.date(), time(LUNCH_END_HOUR, 0))

    if current_time >= time(WORKDAY_END_HOUR, 0):
        next_day = next_workday(current.date())
        return datetime.combine(next_day, time(WORKDAY_START_HOUR, 0))

    return current
########################################################################################################################
def subtract_work_hours(end_dt, hours_to_subtract):
    current = move_to_previous_work_time(end_dt)
    remaining = float(hours_to_subtract)

    if remaining <= 0:
        return current

    while remaining > 0:
        current = move_to_previous_work_time(current)

        if current.time() > time(LUNCH_END_HOUR, 0):
            segment_start = datetime.combine(current.date(), time(LUNCH_END_HOUR, 0))
            available = (current - segment_start).total_seconds() / 3600.0
        elif current.time() > time(WORKDAY_START_HOUR, 0):
            segment_start = datetime.combine(current.date(), time(WORKDAY_START_HOUR, 0))
            available = (current - segment_start).total_seconds() / 3600.0
        else:
            available = 0

        if available <= 0:
            prev_day = previous_workday(current.date())
            current = datetime.combine(prev_day, time(WORKDAY_END_HOUR, 0))
            continue

        if remaining <= available:
            return current - timedelta(hours=remaining)

        remaining -= available
        prev_day = previous_workday(current.date())
        current = datetime.combine(prev_day, time(WORKDAY_END_HOUR, 0))

    return current
########################################################################################################################
def add_work_hours(start_dt, hours_to_add):
    current = move_to_next_work_time(start_dt)
    remaining = float(hours_to_add)

    if remaining <= 0:
        return current

    while remaining > 0:
        current = move_to_next_work_time(current)

        if current.time() < time(LUNCH_START_HOUR, 0):
            segment_end = datetime.combine(current.date(), time(LUNCH_START_HOUR, 0))
            available = (segment_end - current).total_seconds() / 3600.0
        elif current.time() < time(WORKDAY_END_HOUR, 0):
            segment_end = datetime.combine(current.date(), time(WORKDAY_END_HOUR, 0))
            available = (segment_end - current).total_seconds() / 3600.0
        else:
            available = 0

        if available <= 0:
            next_day = next_workday(current.date())
            current = datetime.combine(next_day, time(WORKDAY_START_HOUR, 0))
            continue

        if remaining <= available:
            return current + timedelta(hours=remaining)

        remaining -= available
        if current.time() < time(LUNCH_START_HOUR, 0):
            current = datetime.combine(current.date(), time(LUNCH_END_HOUR, 0))
        else:
            next_day = next_workday(current.date())
            current = datetime.combine(next_day, time(WORKDAY_START_HOUR, 0))

    return current
########################################################################################################################
def working_hours_between(start_dt, end_dt):
    if end_dt <= start_dt:
        return 0.0

    current = move_to_next_work_time(start_dt)
    target = end_dt
    total = 0.0

    while current < target:
        current = move_to_next_work_time(current)

        if current.time() < time(LUNCH_START_HOUR, 0):
            segment_end = datetime.combine(current.date(), time(LUNCH_START_HOUR, 0))
        elif current.time() < time(WORKDAY_END_HOUR, 0):
            segment_end = datetime.combine(current.date(), time(WORKDAY_END_HOUR, 0))
        else:
            next_day = next_workday(current.date())
            current = datetime.combine(next_day, time(WORKDAY_START_HOUR, 0))
            continue

        actual_end = min(segment_end, target)

        if actual_end > current:
            total += (actual_end - current).total_seconds() / 3600.0

        if actual_end >= target:
            break

        if segment_end.time() == time(LUNCH_START_HOUR, 0):
            current = datetime.combine(current.date(), time(LUNCH_END_HOUR, 0))
        else:
            next_day = next_workday(current.date())
            current = datetime.combine(next_day, time(WORKDAY_START_HOUR, 0))

    return total
########################################################################################################################
def working_days_late(required_finish, actual_finish):
    if actual_finish <= required_finish:
        return 0
    hours_late = working_hours_between(required_finish, actual_finish)
    return int(math.ceil(hours_late / HOURS_PER_WORKDAY))
########################################################################################################################
def overlaps(start_a, end_a, start_b, end_b):
    return start_a < end_b and start_b < end_a
########################################################################################################################
def get_resource_bucket_for_task(task_row):
    if task_row.resource_id is not None:
        return f"RESOURCE_{task_row.resource_id}"
    return f"TYPE_{task_row.resource_type_id}"
########################################################################################################################
def build_resource_capacity_map():
    resources = get_resources()
    capacity_map = {}
    by_type_map = {}

    for row in resources:
        member_count = count_resource_members(row.resource_members)
        resource_type_value = int(row.resource_types) if row.resource_types is not None else None

        capacity_map[int(row.resource_id)] = {
            "resource_id": int(row.resource_id),
            "resource_name": row.resource_name,
            "resource_type_id": resource_type_value,
            "default_hours_per_day": float(row.default_hours_per_day) if row.default_hours_per_day is not None else 8.0,
            "resource_members": member_count,
        }

        if resource_type_value is not None:
            by_type_map.setdefault(resource_type_value, []).append(capacity_map[int(row.resource_id)])

    return capacity_map, by_type_map
########################################################################################################################
def get_effective_hours(planned_hours, default_hours_per_day, resource_members):
    if planned_hours is None:
        return 0.0

    planned_hours_value = float(planned_hours)
    members = resource_members if resource_members and resource_members > 0 else 1

    if default_hours_per_day is not None and float(default_hours_per_day) > 7:
        return planned_hours_value / members

    return planned_hours_value
########################################################################################################################
def find_latest_slot_before_deadline(bookings, required_finish, duration_hours, earliest_allowed_start):
    candidate_finish = move_to_previous_work_time(required_finish)

    while True:
        candidate_start = subtract_work_hours(candidate_finish, duration_hours)

        if candidate_start < earliest_allowed_start:
            return None, None

        overlap_found = None
        for booking_start, booking_finish in bookings:
            if overlaps(candidate_start, candidate_finish, booking_start, booking_finish):
                overlap_found = (booking_start, booking_finish)
                break

        if overlap_found is None:
            return candidate_start, candidate_finish

        candidate_finish = move_to_previous_work_time(overlap_found[0])
########################################################################################################################
def find_earliest_slot_after_start(bookings, earliest_allowed_start, duration_hours):
    candidate_start = move_to_next_work_time(earliest_allowed_start)

    while True:
        candidate_finish = add_work_hours(candidate_start, duration_hours)

        overlap_found = None
        for booking_start, booking_finish in bookings:
            if overlaps(candidate_start, candidate_finish, booking_start, booking_finish):
                overlap_found = (booking_start, booking_finish)
                break

        if overlap_found is None:
            return candidate_start, candidate_finish

        candidate_start = move_to_next_work_time(overlap_found[1])
########################################################################################################################
def run_backward_scheduler(entity_type, entity_id):
    if entity_type == "job":
        header = fetch_one(
            """
            SELECT j.job_id AS entity_id, j.due_date, c.customer_importance
            FROM jobs AS j
            LEFT JOIN customers AS c ON j.customer_id = c.customer_id
            WHERE j.job_id = ?
            """,
            [entity_id],
        )
        line_items = fetch_all(
            """
            SELECT line_item_id
            FROM line_items
            WHERE job_number_id = ?
            """,
            [entity_id],
        )
    else:
        header = fetch_one(
            """
            SELECT w.wo_id AS entity_id, w.due_date, c.customer_importance
            FROM work_orders AS w
            LEFT JOIN customers AS c ON w.customer_id = c.customer_id
            WHERE w.wo_id = ?
            """,
            [entity_id],
        )
        line_items = fetch_all(
            """
            SELECT line_item_id
            FROM line_items
            WHERE work_order_id = ?
            """,
            [entity_id],
        )

    if not header:
        return {"scheduled": 0, "late_tasks": 0, "late_final_tasks": 0, "loss": 0.0, "message": "Record not found."}

    due_dt = combine_date_and_workday_end(header.due_date)
    customer_importance = float(header.customer_importance) if header.customer_importance is not None else 1.0

    line_item_ids = [row.line_item_id for row in line_items]
    if not line_item_ids:
        return {"scheduled": 0, "late_tasks": 0, "late_final_tasks": 0, "loss": 0.0, "message": "No line items found."}

    task_records = []
    for line_item_id in line_item_ids:
        cost_codes = fetch_all(
            """
            SELECT cost_code_id
            FROM cost_codes
            WHERE line_item_id = ?
            """,
            [line_item_id],
        )

        for cc in cost_codes:
            tasks = fetch_all(
                """
                SELECT task_id, cost_code_id, sequence_number, planned_hours, resource_type_id, resource_id
                FROM tasks
                WHERE cost_code_id = ?
                  AND sequence_number IS NOT NULL
                ORDER BY sequence_number DESC, task_id DESC
                """,
                [cc.cost_code_id],
            )
            for task in tasks:
                task_records.append(task)

    if not task_records:
        return {"scheduled": 0, "late_tasks": 0, "late_final_tasks": 0, "loss": 0.0, "message": "No tasks found."}

    capacity_map, by_type_map = build_resource_capacity_map()
    resource_bookings = {resource_id: [] for resource_id in capacity_map.keys()}
    current_boundary = move_to_next_work_time(datetime.now())

    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        """
        UPDATE tasks
        SET planned_start_date = NULL,
            planned_finish_date = NULL
        WHERE cost_code_id IN (
            SELECT cost_code_id
            FROM cost_codes
            WHERE line_item_id IN (
                SELECT line_item_id
                FROM line_items
                WHERE """ + ("job_number_id" if entity_type == "job" else "work_order_id") + """ = ?
            )
        )
          AND sequence_number IS NOT NULL
        """,
        entity_id,
    )
    conn.commit()

    late_tasks = 0
    late_final_tasks = 0
    total_loss = 0.0
    scheduled_count = 0

    cost_code_to_tasks = {}
    for task in task_records:
        cost_code_to_tasks.setdefault(task.cost_code_id, []).append(task)

    for cost_code_id, tasks in cost_code_to_tasks.items():
        tasks_sorted_desc = sorted(tasks, key=lambda x: (x.sequence_number if x.sequence_number is not None else 0, x.task_id), reverse=True)
        required_finish = due_dt
        max_sequence = max([(t.sequence_number if t.sequence_number is not None else 0) for t in tasks_sorted_desc], default=0)

        for task in tasks_sorted_desc:
            if task.planned_hours is None or float(task.planned_hours) <= 0:
                continue
            if task.resource_id is None and task.resource_type_id is None:
                continue

            if task.resource_id is not None:
                candidate_resources = [capacity_map.get(int(task.resource_id))] if capacity_map.get(int(task.resource_id)) else []
            else:
                candidate_resources = by_type_map.get(int(task.resource_type_id), [])

            if not candidate_resources:
                continue

            best_choice = None
            for resource_info in candidate_resources:
                effective_hours = get_effective_hours(task.planned_hours, resource_info["default_hours_per_day"], resource_info["resource_members"])
                bookings = resource_bookings.setdefault(resource_info["resource_id"], [])
                start_dt, finish_dt = find_latest_slot_before_deadline(
                    bookings=bookings,
                    required_finish=required_finish,
                    duration_hours=effective_hours,
                    earliest_allowed_start=current_boundary,
                )
                feasible = True
                if start_dt is None or finish_dt is None:
                    feasible = False
                    start_dt, finish_dt = find_earliest_slot_after_start(
                        bookings=bookings,
                        earliest_allowed_start=current_boundary,
                        duration_hours=effective_hours,
                    )

                comparison_key = (
                    0 if feasible else 1,
                    abs((required_finish - finish_dt).total_seconds()) if feasible else 0,
                    finish_dt,
                    resource_info["resource_id"],
                )
                if best_choice is None or comparison_key < best_choice[0]:
                    best_choice = (comparison_key, resource_info, start_dt, finish_dt)

            if best_choice is None:
                continue

            _, chosen_resource, start_dt, finish_dt = best_choice
            bookings = resource_bookings.setdefault(chosen_resource["resource_id"], [])
            bookings.append((start_dt, finish_dt))
            bookings.sort(key=lambda x: x[0])

            cur.execute(
                """
                UPDATE tasks
                SET planned_start_date = ?,
                    planned_finish_date = ?,
                    resource_id = ?
                WHERE task_id = ?
                """,
                start_dt,
                finish_dt,
                chosen_resource["resource_id"],
                task.task_id,
            )

            days_late = working_days_late(required_finish, finish_dt)
            is_final_task = (task.sequence_number == max_sequence)

            if days_late > 0:
                late_tasks += 1
                if is_final_task:
                    late_final_tasks += 1
                    total_loss += customer_importance * 5.0 * days_late
                else:
                    total_loss += customer_importance * 1.0 * days_late

            required_finish = start_dt
            scheduled_count += 1

    conn.commit()
    conn.close()

    return {
        "scheduled": scheduled_count,
        "late_tasks": late_tasks,
        "late_final_tasks": late_final_tasks,
        "loss": total_loss,
        "message": "Scheduling completed."
    }
########################################################################################################################
def _work_hours_remaining_in_day(dt_value):
    """
    Working hours remaining in the current manufacturing day after dt_value.
    Lunch is excluded.
    """
    current = move_to_next_work_time(dt_value)

    end_of_day = datetime.combine(
        current.date(),
        time(WORKDAY_END_HOUR, 0),
    )

    lunch_start = datetime.combine(
        current.date(),
        time(LUNCH_START_HOUR, 0),
    )

    lunch_end = datetime.combine(
        current.date(),
        time(LUNCH_END_HOUR, 0),
    )

    if current >= end_of_day:
        return 0.0

    if current < lunch_start:
        return (
            (lunch_start - current).total_seconds() / 3600.0
            + (end_of_day - lunch_end).total_seconds() / 3600.0
        )

    return (
        end_of_day - current
    ).total_seconds() / 3600.0
########################################################################################################################
def compute_inter_task_slack_end(finish_dt):
    """
    Earliest start for the next sequential manufacturing task.

    This reproduces the older global-scheduler behavior:
    use the larger of:
        - the remaining working hours in the current day
        - four working hours
    """
    finish = move_to_next_work_time(finish_dt)

    buffer_hours = max(
        _work_hours_remaining_in_day(finish),
        4.0,
    )

    return add_work_hours(
        finish,
        buffer_hours,
    )
########################################################################################################################
def _safe_manufacturing_due_date(entity_type, entity_id):
    """
    Return manufacturing_due_date when that field exists and is populated.
    Otherwise fall back to due_date.

    The try/fallback keeps this compatible if work_orders does not yet have
    manufacturing_due_date in the Access database.
    """

    if entity_type == "job":

        try:
            row = fetch_one(
                """
                SELECT
                    manufacturing_due_date,
                    due_date
                FROM jobs
                WHERE job_id = ?
                """,
                [entity_id],
            )

            if row:
                return (
                    row.manufacturing_due_date
                    or row.due_date
                )

        except Exception:
            row = fetch_one(
                """
                SELECT due_date
                FROM jobs
                WHERE job_id = ?
                """,
                [entity_id],
            )

            return (
                row.due_date
                if row
                else None
            )

    else:

        try:
            row = fetch_one(
                """
                SELECT
                    manufacturing_due_date,
                    due_date
                FROM work_orders
                WHERE wo_id = ?
                """,
                [entity_id],
            )

            if row:
                return (
                    row.manufacturing_due_date
                    or row.due_date
                )

        except Exception:
            row = fetch_one(
                """
                SELECT due_date
                FROM work_orders
                WHERE wo_id = ?
                """,
                [entity_id],
            )

            return (
                row.due_date
                if row
                else None
            )

    return None
########################################################################################################################
def _mark_manufacturing_scheduled(entity_type, entity_id):
    """
    Mark the Job / WO as manufacturing scheduled when the field exists.
    """
    try:

        if entity_type == "job":

            execute_sql(
                """
                UPDATE jobs
                SET manufacturing_scheduled = ?
                WHERE job_id = ?
                """,
                [
                    -1,
                    entity_id,
                ],
            )

        elif entity_type == "wo":

            execute_sql(
                """
                UPDATE work_orders
                SET manufacturing_scheduled = ?
                WHERE wo_id = ?
                """,
                [
                    -1,
                    entity_id,
                ],
            )

    except Exception:
        # Keep scheduling functional even if an older DB is missing the flag.
        pass
########################################################################################################################
def run_global_scheduler():
    """
    Re-schedule all movable MANUFACTURING tasks across all Jobs and WOs.

    Core behavior:
      - schedules the whole manufacturing shop globally
      - completed / started tasks are fixed and are never moved
      - tasks within each cost code obey sequence_number precedence
      - all ready task chains compete for shared resource capacity
      - priority is:
            customer importance
            x final-task penalty
            x (1 + estimated working days late)
      - manufacturing_due_date is used when available, otherwise due_date
      - if no specific resource is locked, the resource that finishes earliest wins
      - crew size is respected through get_effective_hours()
      - normal sequential tasks receive inter-task handoff slack
      - special-order tasks receive a four-working-day gap when that field exists
      - Installation / Survey resource type 13 is NOT touched
      - Tech Ops / Overhead indirect task IDs are NOT touched
    """

    now = datetime.now()
    now_work = move_to_next_work_time(now)

    complete_sid = (
        get_status_id_by_name("complete")
        or get_status_id_by_name("Complete")
    )

    # Any task that has ever been clocked into is fixed.
    punched_ids = {
        int(row.task_id)
        for row in fetch_all(
            """
            SELECT DISTINCT task_id
            FROM time_punches
            WHERE task_id IS NOT NULL
            """
        )
        if row.task_id is not None
    }

    # -----------------------------------------------------------------
    # BUILD JOB / WO DUE-DATE AND CUSTOMER-IMPORTANCE MAP
    # -----------------------------------------------------------------

    entity_map = {}

    job_rows = fetch_all(
        """
        SELECT
            j.job_id,
            j.due_date,
            j.customer_id
        FROM jobs AS j
        """
    )

    for job in job_rows:

        importance = 1.0

        if job.customer_id:

            customer = fetch_one(
                """
                SELECT customer_importance
                FROM customers
                WHERE customer_id = ?
                """,
                [job.customer_id],
            )

            if (
                customer
                and customer.customer_importance is not None
            ):
                importance = float(
                    customer.customer_importance
                )

        raw_due = _safe_manufacturing_due_date(
            "job",
            job.job_id,
        )

        entity_map[
            ("job", int(job.job_id))
        ] = {
            "due_dt": (
                combine_date_and_workday_end(raw_due)
                if raw_due
                else None
            ),
            "importance": importance,
        }

    work_order_rows = fetch_all(
        """
        SELECT
            w.wo_id,
            w.due_date,
            w.customer_id
        FROM work_orders AS w
        """
    )

    for wo in work_order_rows:

        importance = 1.0

        if wo.customer_id:

            customer = fetch_one(
                """
                SELECT customer_importance
                FROM customers
                WHERE customer_id = ?
                """,
                [wo.customer_id],
            )

            if (
                customer
                and customer.customer_importance is not None
            ):
                importance = float(
                    customer.customer_importance
                )

        raw_due = _safe_manufacturing_due_date(
            "wo",
            wo.wo_id,
        )

        entity_map[
            ("wo", int(wo.wo_id))
        ] = {
            "due_dt": (
                combine_date_and_workday_end(raw_due)
                if raw_due
                else None
            ),
            "importance": importance,
        }

    # -----------------------------------------------------------------
    # RESOURCE CAPACITY
    # -----------------------------------------------------------------

    capacity_map, by_type_map = (
        build_resource_capacity_map()
    )

    # Manufacturing scheduler should never assign Installation resources.
    capacity_map = {
        rid: info
        for rid, info in capacity_map.items()
        if int(
            info.get(
                "resource_type_id"
            )
            or 0
        ) != INSTALLATION_RESOURCE_TYPE_ID
    }

    by_type_map = {
        rt_id: [
            info
            for info in rows
            if int(
                info.get(
                    "resource_type_id"
                )
                or 0
            ) != INSTALLATION_RESOURCE_TYPE_ID
        ]
        for rt_id, rows in by_type_map.items()
        if int(rt_id or 0)
        != INSTALLATION_RESOURCE_TYPE_ID
    }

    resource_bookings = {
        rid: []
        for rid in capacity_map
    }

    # -----------------------------------------------------------------
    # BUILD TASK GROUPS
    #
    # Current app stores manufacturing tasks by cost_code_id for both
    # Jobs and Work Orders.
    # -----------------------------------------------------------------

    groups = {}

    task_rows = fetch_all(
        """
        SELECT
            task_id,
            cost_code_id,
            sequence_number,
            planned_hours,
            resource_type_id,
            resource_id,
            status_id,
            percentage_complete,
            planned_start_date,
            planned_finish_date
        FROM tasks
        WHERE cost_code_id IS NOT NULL
          AND cost_code_id > 0
          AND sequence_number IS NOT NULL
          AND planned_hours IS NOT NULL
          AND planned_hours > 0
        ORDER BY
            cost_code_id,
            sequence_number,
            task_id
        """
    )

    for task in task_rows:

        if (
            task.task_id in INDIRECT_TASK_IDS
        ):
            continue

        if (
            task.resource_type_id is not None
            and int(task.resource_type_id)
            == INSTALLATION_RESOURCE_TYPE_ID
        ):
            continue

        cost_code = fetch_one(
            """
            SELECT
                cost_code_id,
                line_item_id
            FROM cost_codes
            WHERE cost_code_id = ?
            """,
            [task.cost_code_id],
        )

        if (
            not cost_code
            or not cost_code.line_item_id
        ):
            continue

        line_item = fetch_one(
            """
            SELECT
                line_item_id,
                job_number_id,
                work_order_id
            FROM line_items
            WHERE line_item_id = ?
            """,
            [cost_code.line_item_id],
        )

        if not line_item:
            continue

        if line_item.job_number_id is not None:

            entity_type = "job"
            entity_id = int(
                line_item.job_number_id
            )

        elif line_item.work_order_id is not None:

            entity_type = "wo"
            entity_id = int(
                line_item.work_order_id
            )

        else:
            continue

        group_key = (
            "cc",
            int(task.cost_code_id),
        )

        if group_key not in groups:

            groups[group_key] = {
                "entity_type": entity_type,
                "entity_id": entity_id,
                "tasks": [],
            }

        groups[group_key]["tasks"].append(
            task
        )

    # -----------------------------------------------------------------
    # OPTIONAL SPECIAL-ORDER FLAGS
    # -----------------------------------------------------------------

    special_order_ids = set()

    try:

        special_rows = fetch_all(
            """
            SELECT task_id
            FROM tasks
            WHERE special_order = -1
            """
        )

        special_order_ids = {
            int(row.task_id)
            for row in special_rows
        }

    except Exception:
        # Current database does not need this field for the scheduler to work.
        pass

    # -----------------------------------------------------------------
    # CLASSIFY TASKS AS FIXED OR MOVABLE
    # -----------------------------------------------------------------

    for group in groups.values():

        raw_tasks = group["tasks"]

        raw_tasks.sort(
            key=lambda task: (
                task.sequence_number or 0,
                task.task_id,
            )
        )

        max_sequence = max(
            [
                task.sequence_number or 0
                for task in raw_tasks
            ],
            default=0,
        )

        converted = []

        for task in raw_tasks:

            is_complete = (
                (
                    complete_sid is not None
                    and task.status_id
                    == complete_sid
                )
                or float(
                    task.percentage_complete
                    or 0
                ) >= 100.0
            )

            is_fixed = (
                is_complete
                or int(task.task_id)
                in punched_ids
            )

            converted.append(
                {
                    "task_id":
                        int(task.task_id),

                    "seq":
                        int(
                            task.sequence_number
                            or 0
                        ),

                    "hours":
                        float(
                            task.planned_hours
                            or 0
                        ),

                    "resource_type_id":
                        (
                            int(task.resource_type_id)
                            if task.resource_type_id
                            else None
                        ),

                    "resource_id":
                        (
                            int(task.resource_id)
                            if (
                                task.resource_id
                                and int(
                                    task.resource_id
                                ) > 0
                            )
                            else None
                        ),

                    "is_fixed":
                        is_fixed,

                    "is_final":
                        (
                            int(
                                task.sequence_number
                                or 0
                            )
                            == int(
                                max_sequence
                                or 0
                            )
                        ),

                    "special_order":
                        (
                            int(task.task_id)
                            in special_order_ids
                        ),

                    "sched_start":
                        (
                            task.planned_start_date
                            if is_fixed
                            else None
                        ),

                    "sched_finish":
                        (
                            task.planned_finish_date
                            if is_fixed
                            else None
                        ),

                    "assigned_rid":
                        (
                            int(task.resource_id)
                            if (
                                is_fixed
                                and task.resource_id
                                and int(
                                    task.resource_id
                                ) > 0
                            )
                            else None
                        ),

                    "dispatched":
                        is_fixed,
                }
            )

        group["tasks"] = converted

    # -----------------------------------------------------------------
    # FIXED TASKS OCCUPY RESOURCE CAPACITY
    # -----------------------------------------------------------------

    for group in groups.values():

        for task in group["tasks"]:

            if (
                task["is_fixed"]
                and task["assigned_rid"]
                and task["sched_start"]
                and task["sched_finish"]
            ):

                rid = task["assigned_rid"]

                if rid in resource_bookings:

                    resource_bookings[
                        rid
                    ].append(
                        (
                            task["sched_start"],
                            task["sched_finish"],
                        )
                    )

    for rid in resource_bookings:

        resource_bookings[rid].sort(
            key=lambda row: row[0]
        )

    # -----------------------------------------------------------------
    # INITIAL EARLIEST START FOR EACH TASK CHAIN
    # -----------------------------------------------------------------

    for group in groups.values():

        earliest = now_work

        for task in group["tasks"]:

            if not task["is_fixed"]:
                break

            finish = (
                task["sched_finish"]
                or now_work
            )

            if task["special_order"]:

                gap_end = add_work_hours(
                    finish,
                    SPECIAL_ORDER_GAP_DAYS
                    * HOURS_PER_WORKDAY,
                )

                earliest = max(
                    earliest,
                    gap_end,
                )

            else:

                earliest = max(
                    earliest,
                    compute_inter_task_slack_end(
                        finish
                    ),
                )

        group["earliest_next"] = earliest

    # -----------------------------------------------------------------
    # GLOBAL PRIORITY-DISPATCH LOOP
    # -----------------------------------------------------------------

    total_task_count = sum(
        len(group["tasks"])
        for group in groups.values()
    )

    max_iterations = (
        total_task_count
        + 10
    )

    for _ in range(max_iterations):

        ready = []

        for group_key, group in groups.items():

            for task in group["tasks"]:

                if task["dispatched"]:
                    continue

                predecessors_done = all(
                    predecessor["dispatched"]
                    for predecessor
                    in group["tasks"]
                    if (
                        predecessor["seq"]
                        < task["seq"]
                    )
                )

                if predecessors_done:

                    ready.append(
                        (
                            group_key,
                            group,
                            task,
                        )
                    )

                # Only the first undispatched task
                # in each chain can be ready.
                break

        if not ready:
            break

        def priority_key(item):

            _group_key, group, task = item

            entity_key = (
                group["entity_type"],
                group["entity_id"],
            )

            entity_info = (
                entity_map.get(
                    entity_key,
                    {},
                )
            )

            due_date = entity_info.get(
                "due_dt"
            )

            importance = float(
                entity_info.get(
                    "importance",
                    1.0,
                )
            )

            penalty = (
                5.0
                if task["is_final"]
                else 1.0
            )

            estimated_finish = add_work_hours(
                group["earliest_next"],
                task["hours"],
            )

            estimated_days_late = (
                working_days_late(
                    due_date,
                    estimated_finish,
                )
                if due_date
                else 0
            )

            due_timestamp = (
                due_date.timestamp()
                if due_date
                else 1e18
            )

            weighted_priority = (
                importance
                * penalty
                * (
                    1.0
                    + estimated_days_late
                )
            )

            # Python sorts ascending, so negate
            # the weighted score to make highest priority first.
            return (
                -weighted_priority,
                due_timestamp,
                task["task_id"],
            )

        ready.sort(
            key=priority_key
        )

        _group_key, group, task = (
            ready[0]
        )

        # -------------------------------------------------------------
        # RESOURCE CANDIDATES
        # -------------------------------------------------------------

        if task["resource_id"]:

            resource = capacity_map.get(
                task["resource_id"]
            )

            candidates = (
                [resource]
                if resource
                else []
            )

        elif task["resource_type_id"]:

            candidates = by_type_map.get(
                task["resource_type_id"],
                [],
            )

        else:

            candidates = []

        if not candidates:

            # Mark dispatched so this missing-resource task
            # does not block the entire global scheduler.
            task["dispatched"] = True
            continue

        # -------------------------------------------------------------
        # CHOOSE RESOURCE WITH EARLIEST FINISH
        # -------------------------------------------------------------

        best_start = None
        best_finish = None
        best_resource = None

        for resource in candidates:

            effective_hours = get_effective_hours(
                task["hours"],
                resource[
                    "default_hours_per_day"
                ],
                resource[
                    "resource_members"
                ],
            )

            bookings = (
                resource_bookings.get(
                    resource["resource_id"],
                    [],
                )
            )

            start_dt, finish_dt = (
                find_earliest_slot_after_start(
                    bookings=bookings,
                    earliest_allowed_start=
                        group[
                            "earliest_next"
                        ],
                    duration_hours=
                        effective_hours,
                )
            )

            if (
                best_finish is None
                or finish_dt < best_finish
            ):

                best_start = start_dt
                best_finish = finish_dt
                best_resource = resource

        task["sched_start"] = (
            best_start
        )

        task["sched_finish"] = (
            best_finish
        )

        task["assigned_rid"] = (
            best_resource[
                "resource_id"
            ]
            if best_resource
            else None
        )

        task["dispatched"] = True

        if best_resource:

            rid = best_resource[
                "resource_id"
            ]

            resource_bookings.setdefault(
                rid,
                [],
            ).append(
                (
                    best_start,
                    best_finish,
                )
            )

            resource_bookings[
                rid
            ].sort(
                key=lambda row: row[0]
            )

        # -------------------------------------------------------------
        # ADVANCE THIS TASK CHAIN
        # -------------------------------------------------------------

        if best_finish:

            if task["special_order"]:

                group["earliest_next"] = (
                    add_work_hours(
                        best_finish,
                        SPECIAL_ORDER_GAP_DAYS
                        * HOURS_PER_WORKDAY,
                    )
                )

            else:

                group["earliest_next"] = (
                    compute_inter_task_slack_end(
                        best_finish
                    )
                )

        # If following tasks are already fixed,
        # fast-forward through them.
        current_index = (
            group["tasks"].index(task)
        )

        for next_task in (
            group["tasks"][
                current_index + 1:
            ]
        ):

            if not next_task["is_fixed"]:
                break

            if not next_task["sched_finish"]:
                continue

            if next_task["special_order"]:

                gap_end = add_work_hours(
                    next_task["sched_finish"],
                    SPECIAL_ORDER_GAP_DAYS
                    * HOURS_PER_WORKDAY,
                )

                group["earliest_next"] = max(
                    group["earliest_next"],
                    gap_end,
                )

            else:

                handoff_end = (
                    compute_inter_task_slack_end(
                        next_task[
                            "sched_finish"
                        ]
                    )
                )

                group["earliest_next"] = max(
                    group["earliest_next"],
                    handoff_end,
                )

    # -----------------------------------------------------------------
    # WRITE NEW GLOBAL SCHEDULE
    # -----------------------------------------------------------------

    conn = get_connection()
    cur = conn.cursor()

    # Clear only movable manufacturing tasks.
    for group in groups.values():

        for task in group["tasks"]:

            if task["is_fixed"]:
                continue

            cur.execute(
                """
                UPDATE tasks
                SET
                    planned_start_date = NULL,
                    planned_finish_date = NULL
                WHERE task_id = ?
                """,
                task["task_id"],
            )

    conn.commit()

    scheduled_count = 0
    skipped_count = 0

    for group in groups.values():

        for task in group["tasks"]:

            if task["is_fixed"]:
                continue

            if (
                task["sched_start"]
                and task["sched_finish"]
                and task["assigned_rid"]
            ):

                cur.execute(
                    """
                    UPDATE tasks
                    SET
                        planned_start_date = ?,
                        planned_finish_date = ?,
                        resource_id = ?
                    WHERE task_id = ?
                    """,
                    task["sched_start"],
                    task["sched_finish"],
                    task["assigned_rid"],
                    task["task_id"],
                )

                scheduled_count += 1

            else:

                skipped_count += 1

    conn.commit()
    conn.close()

    # -----------------------------------------------------------------
    # CALCULATE RESULTING WEIGHTED LATENESS LOSS
    # -----------------------------------------------------------------

    late_tasks = 0
    late_final_tasks = 0
    total_loss = 0.0

    for group in groups.values():

        entity_key = (
            group["entity_type"],
            group["entity_id"],
        )

        entity_info = entity_map.get(
            entity_key,
            {},
        )

        due_date = entity_info.get(
            "due_dt"
        )

        importance = float(
            entity_info.get(
                "importance",
                1.0,
            )
        )

        if not due_date:
            continue

        for task in group["tasks"]:

            finish = task[
                "sched_finish"
            ]

            if not finish:
                continue

            days_late = (
                working_days_late(
                    due_date,
                    finish,
                )
            )

            if days_late <= 0:
                continue

            penalty = (
                5.0
                if task["is_final"]
                else 1.0
            )

            late_tasks += 1

            if task["is_final"]:
                late_final_tasks += 1

            total_loss += (
                importance
                * penalty
                * days_late
            )

    return {
        "scheduled": scheduled_count,
        "skipped": skipped_count,
        "late_tasks": late_tasks,
        "late_final_tasks": late_final_tasks,
        "loss": total_loss,
        "message": "Global manufacturing schedule optimized.",
    }
########################################################################################################################
########################################################################################################################
def get_home_gantt_data(weeks=6):
    today = datetime.now().date()
    week_start = today - timedelta(days=today.weekday())
    horizon_end = week_start + timedelta(days=(weeks * 7) - 1)

    visible_days = []
    current_day = week_start
    while current_day <= horizon_end:
        if current_day.weekday() < 5:
            visible_days.append(current_day)
        current_day += timedelta(days=1)

    hour_labels = ["7a", "8a", "9a", "10a", "11a", "12p", "1p", "2p", "3p"]
    day_index_map = {day: idx for idx, day in enumerate(visible_days)}

    tasks = fetch_all(
        """
        SELECT 
            task_id,
            task_name,
            task_description,
            sequence_number,
            planned_hours,
            planned_start_date,
            planned_finish_date,
            resource_id,
            resource_type_id,
            cost_code_id,
            status_id
        FROM tasks
        WHERE planned_start_date IS NOT NULL
          AND planned_finish_date IS NOT NULL
          AND resource_id IS NOT NULL
        ORDER BY planned_start_date, planned_finish_date, task_id
        """
    )

    all_resources = fetch_all(
        """
        SELECT resource_id, resource_name
        FROM resources
        ORDER BY resource_name
        """
    )

    rows_by_resource = {}
    for resource in all_resources:
        rows_by_resource[resource.resource_name] = []

    resource_name_map = {r.resource_id: r.resource_name for r in all_resources}
    resource_type_name_map = get_resource_type_name_map()

    def get_employee_status_label(status_id):
        status_value = int(status_id or 0)

        if status_value == 1:
            return "NOT STARTED"
        elif status_value == 2:
            return "STARTED"
        elif status_value == 3:
            return "COMPLETE"
        else:
            return "UNKNOWN"

    def position_from_datetime(dt_value):
        day_value = dt_value.date()
        if day_value not in day_index_map:
            return None

        minutes = dt_value.hour * 60 + dt_value.minute

        if minutes <= 7 * 60:
            hour_offset = 0.0
        elif minutes < 12 * 60:
            hour_offset = (minutes - 7 * 60) / 60.0
        elif minutes < 13 * 60:
            hour_offset = 5.0
        elif minutes <= 16 * 60:
            hour_offset = 6.0 + (minutes - 13 * 60) / 60.0
        else:
            hour_offset = 9.0

        return day_index_map[day_value] * 9.0 + hour_offset

    for row in tasks:
        start_dt = row.planned_start_date
        finish_dt = row.planned_finish_date

        if finish_dt < datetime.combine(week_start, time(0, 0)):
            continue
        if start_dt > datetime.combine(horizon_end, time(23, 59)):
            continue
        if start_dt.date() not in day_index_map or finish_dt.date() not in day_index_map:
            continue

        start_pos = position_from_datetime(start_dt)
        finish_pos = position_from_datetime(finish_dt)

        if start_pos is None or finish_pos is None:
            continue

        if finish_pos <= start_pos:
            finish_pos = start_pos + 0.25

        resource_name = resource_name_map.get(row.resource_id, f"Resource {row.resource_id}")
        resource_type_name = resource_type_name_map.get(row.resource_type_id, "")
        status_label = get_employee_status_label(row.status_id)

        if resource_name not in rows_by_resource:
            rows_by_resource[resource_name] = []

        rows_by_resource[resource_name].append({
            "task_id": row.task_id,
            "task_name": row.task_name or f"Task {row.task_id}",
            "task_description": row.task_description or "",
            "sequence_number": row.sequence_number or "",
            "planned_hours": row.planned_hours or "",
            "resource_name": resource_name,
            "resource_type": resource_type_name,
            "cost_code_id": row.cost_code_id or "",
            "status": status_label,
            "start_label": start_dt.strftime('%Y-%m-%d %I:%M %p'),
            "finish_label": finish_dt.strftime('%Y-%m-%d %I:%M %p'),
            "start_pos": start_pos,
            "span_pos": max(0.25, finish_pos - start_pos),
        })

    gantt_rows = []
    for resource in all_resources:
        resource_name = resource.resource_name
        gantt_rows.append({
            "resource_name": resource_name,
            "tasks": sorted(
                rows_by_resource.get(resource_name, []),
                key=lambda x: (x["start_pos"], x["task_id"])
            ),
        })

    day_columns = []
    for day in visible_days:
        day_columns.append({
            "label_top": day.strftime('%a'),
            "label_bottom": day.strftime('%m/%d'),
            "hours": hour_labels,
        })

    return {
        "week_start_label":  week_start.strftime("%Y-%m-%d"),
        "horizon_end_label": horizon_end.strftime("%Y-%m-%d"),
        "day_columns":       day_columns,
        "gantt_rows":        gantt_rows,
        "total_hour_columns": len(visible_days) * 9,
    }
########################################################################################################################
def get_home_task_table_data():
    tasks = fetch_all(
        """
        SELECT 
            t.task_id,
            t.task_name,
            t.task_description,
            t.planned_hours,
            t.resource_type_id,
            t.resource_id,
            t.planned_start_date,
            t.planned_finish_date,
            t.status_id,
            t.percentage_complete,
            
            cc.coins_cost_code,

            li.job_number_id,
            li.work_order_id,

            ej.employee_first_name AS job_pm_first_name,
            ej.employee_last_name AS job_pm_last_name,

            ew.employee_first_name AS wo_pm_first_name,
            ew.employee_last_name AS wo_pm_last_name,

            SUM(tp.task_time) AS current_minutes

        FROM 
            ((((((tasks AS t
            LEFT JOIN time_punches AS tp
                ON t.task_id = tp.task_id)
            LEFT JOIN cost_codes AS cc
                ON t.cost_code_id = cc.cost_code_id)
            LEFT JOIN line_items AS li
                ON cc.line_item_id = li.line_item_id)
            LEFT JOIN jobs AS j
                ON li.job_number_id = j.job_id)
            LEFT JOIN work_orders AS w
                ON li.work_order_id = w.wo_id)
            LEFT JOIN employees AS ej
                ON j.project_manager_employee_id = ej.employee_id)
            LEFT JOIN employees AS ew
                ON w.project_manager_employee_id = ew.employee_id

        GROUP BY
            t.task_id,
            t.task_name,
            t.task_description,
            t.planned_hours,
            t.resource_type_id,
            t.resource_id,
            t.planned_start_date,
            t.planned_finish_date,
            t.status_id,
            t.percentage_complete,
            
            cc.coins_cost_code,

            li.job_number_id,
            li.work_order_id,

            ej.employee_first_name,
            ej.employee_last_name,

            ew.employee_first_name,
            ew.employee_last_name

        ORDER BY t.planned_start_date, t.task_id
        """
    )

    resource_name_map = get_resource_name_map()
    resource_type_name_map = get_resource_type_name_map()

    rows = []

    for row in tasks:
        current_minutes = float(row.current_minutes or 0)
        current_hours = current_minutes / 60.0

        if int(row.status_id or 0) == 1:
            status_label = "NOT STARTED"
        elif int(row.status_id or 0) == 2:
            status_label = "STARTED"
        else:
            status_label = "UNKNOWN"

        if row.job_number_id:
            job_wo_number = f"Job {row.job_number_id}"
            project_manager = f"{row.job_pm_first_name or ''} {row.job_pm_last_name or ''}".strip()
        elif row.work_order_id:
            job_wo_number = f"WO {row.work_order_id}"
            project_manager = f"{row.wo_pm_first_name or ''} {row.wo_pm_last_name or ''}".strip()
        else:
            job_wo_number = ""
            project_manager = ""

        rows.append({
            "cost_code": row.coins_cost_code or "",
            "task_id": row.task_id,
            "job_wo_number": job_wo_number,
            "project_manager": project_manager,
            "task_name": row.task_name or "",
            "task_description": row.task_description or "",
            "planned_hours": row.planned_hours or 0,
            "current_hours": current_hours,
            "percent_complete": row.percentage_complete or 0,
            "resource_type": resource_type_name_map.get(row.resource_type_id, "") if row.resource_type_id is not None else "",
            "resource": resource_name_map.get(row.resource_id, "") if row.resource_id is not None else "",
            "planned_start": row.planned_start_date or "",
            "planned_finish": row.planned_finish_date or "",
            "status": status_label,
        })

    return rows
########################################################################################################################
def update_task_percentage_complete(task_id):
    task = fetch_one(
        """
        SELECT task_id, planned_hours
        FROM tasks
        WHERE task_id = ?
        """,
        [task_id],
    )

    if not task:
        return

    planned_hours = float(task.planned_hours or 0)
    planned_minutes = planned_hours * 60

    time_row = fetch_one(
        """
        SELECT SUM(task_time) AS total_minutes
        FROM time_punches
        WHERE task_id = ?
        """,
        [task_id],
    )

    total_minutes = float(time_row.total_minutes or 0)

    if planned_minutes > 0:
        percentage_complete = round((total_minutes / planned_minutes) * 100, 1)
    else:
        percentage_complete = 0

    execute_sql(
        """
        UPDATE tasks
        SET percentage_complete = ?
        WHERE task_id = ?
        """,
        [percentage_complete, task_id],
    )
########################################################################################################################
########################################################################################################################
########################################################################################################################





########################################################################################################################
########################################################################################################################
########################################################################################################################
"""INSTALLATION SCHEDULING HELPERS"""
########################################################################################################################
########################################################################################################################
########################################################################################################################
INSTALLATION_DAILY_CAPACITY_HOURS = 8
INSTALLATION_RESOURCE_TYPE_ID = 13
########################################################################################################################
def _installation_resources():
    return fetch_all(
        """
        SELECT resource_id, resource_name
        FROM resources
        WHERE resource_types = ?
        ORDER BY resource_name
        """,
        [str(INSTALLATION_RESOURCE_TYPE_ID)],
    )
########################################################################################################################
def _installation_qc_forms():
    return fetch_all(
        """
        SELECT
            qc_form_id,
            form_name
        FROM qc_forms
        WHERE resource_type_id = ?
          AND active = True
          AND qc_form_id <> ?
        ORDER BY
            form_name,
            qc_form_id
        """,
        [
            INSTALLATION_RESOURCE_TYPE_ID,
            2,
        ],
    )
########################################################################################################################
def _first_job_cost_code(job_id):
    return fetch_one(
        """
        SELECT TOP 1 cc.cost_code_id
        FROM line_items AS li
        INNER JOIN cost_codes AS cc
            ON cc.line_item_id = li.line_item_id
        WHERE li.job_number_id = ?
        ORDER BY cc.cost_code_id
        """,
        [job_id],
    )
########################################################################################################################
def _installation_completion_form_id():
    row = fetch_one(
        """
        SELECT TOP 1 qc_form_id
        FROM qc_forms
        WHERE resource_type_id = ?
          AND active = True
          AND form_name LIKE '*Installation*'
        ORDER BY qc_form_id
        """,
        [INSTALLATION_RESOURCE_TYPE_ID],
    )
    return row.qc_form_id if row else None
########################################################################################################################
def _job_rt13_tasks(job_id):
    """
    Get all Installation / Survey tasks for one job.

    Uses simple queries instead of a large Access JOIN.
    """

    # ---------------------------------------------------------
    # Get all line items for this job
    # ---------------------------------------------------------
    line_items = fetch_all(
        """
        SELECT line_item_id
        FROM line_items
        WHERE job_number_id = ?
        """,
        [job_id],
    )

    if not line_items:
        return []


    # ---------------------------------------------------------
    # Get all cost codes belonging to those line items
    # ---------------------------------------------------------
    cost_code_ids = []

    for line_item in line_items:
        cost_codes = fetch_all(
            """
            SELECT cost_code_id
            FROM cost_codes
            WHERE line_item_id = ?
            """,
            [line_item.line_item_id],
        )

        for cost_code in cost_codes:
            cost_code_ids.append(cost_code.cost_code_id)


    if not cost_code_ids:
        return []


    # ---------------------------------------------------------
    # Get Installation / Survey tasks
    # ---------------------------------------------------------
    tasks = []

    for cost_code_id in cost_code_ids:
        rows = fetch_all(
            """
            SELECT
                task_id,
                task_name,
                planned_start_date,
                planned_finish_date,
                planned_hours,
                resource_id
            FROM tasks
            WHERE cost_code_id = ?
              AND resource_type_id = ?
            ORDER BY planned_start_date, task_id
            """,
            [
                cost_code_id,
                INSTALLATION_RESOURCE_TYPE_ID,
            ],
        )

        tasks.extend(rows)


    # ---------------------------------------------------------
    # Get resource names separately
    # ---------------------------------------------------------
    resource_name_map = {
        row.resource_id: row.resource_name
        for row in fetch_all(
            """
            SELECT resource_id, resource_name
            FROM resources
            """
        )
    }


    # pyodbc.Row cannot have attributes added to it,
    # so convert the results into simple objects.
    from types import SimpleNamespace

    result = []

    for task in tasks:
        result.append(
            SimpleNamespace(
                task_id=task.task_id,
                task_name=task.task_name,
                planned_start_date=task.planned_start_date,
                planned_finish_date=task.planned_finish_date,
                planned_hours=task.planned_hours,
                resource_id=task.resource_id,
                resource_name=resource_name_map.get(
                    task.resource_id,
                    ""
                ),
            )
        )

    return result
########################################################################################################################
def get_qc_form_for_task(task):
    """
    Survey tasks use tasks.survey_type as the exact qc_form_id.
    Normal tasks fall back to the active QC form for their resource type.
    """

    if getattr(task, "survey_type", None):
        return fetch_one(
            """
            SELECT
                qc_form_id,
                form_name
            FROM qc_forms
            WHERE qc_form_id = ?
              AND active = True
            """,
            [task.survey_type],
        )

    return fetch_one(
        """
        SELECT TOP 1
            qc_form_id,
            form_name
        FROM qc_forms
        WHERE resource_type_id = ?
          AND active = True
        ORDER BY qc_form_id
        """,
        [task.resource_type_id],
    )
########################################################################################################################
def _survey_cost_code_options(job_id):
    """
    Return only cost codes belonging to line items for this job.
    """

    return fetch_all(
        """
        SELECT
            cc.cost_code_id,
            cc.coins_cost_code,
            cc.cost_code_description,
            li.line_item_id,
            li.line_item_description
        FROM line_items AS li
        INNER JOIN cost_codes AS cc
            ON cc.line_item_id = li.line_item_id
        WHERE li.job_number_id = ?
        ORDER BY
            li.line_item_id,
            cc.coins_cost_code,
            cc.cost_code_description,
            cc.cost_code_id
        """,
        [job_id],
    )
########################################################################################################################
########################################################################################################################
########################################################################################################################





########################################################################################################################
########################################################################################################################
########################################################################################################################
"""RESOURCE MANAGEMENT HELPERS"""
########################################################################################################################
########################################################################################################################
########################################################################################################################
def _can_manage_resources(employee):
    """Resource Management access is controlled exclusively by employees.admin."""
    return is_admin_employee(employee)
########################################################################################################################
def _access_id_key(value):
    """
    Normalize Access ID values in Python.

    This avoids Access comparison/type issues if related ID fields are returned
    by pyodbc using different numeric representations.
    """
    if value is None:
        return None

    raw = str(value).strip()

    if raw == "":
        return None

    try:
        return str(int(float(raw)))
    except (TypeError, ValueError):
        return raw
########################################################################################################################
def _parse_int(value, default=0):
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return default
########################################################################################################################
def _parse_float(value, default=0.0):
    try:
        return float(str(value).strip())
    except (TypeError, ValueError):
        return default
########################################################################################################################
def _sync_resource_members_from_employees():
    """
    Keep resources.resource_members synchronized with employees.resource_id.

    employees.resource_id is the source of truth. resource_members remains
    populated because the scheduling engine uses the comma-separated entries
    to determine the number of people assigned to a resource.
    """

    resources = fetch_all(
        """
        SELECT resource_id
        FROM resources
        ORDER BY resource_id
        """
    )

    employees = fetch_all(
        """
        SELECT employee_id, resource_id
        FROM employees
        ORDER BY employee_id
        """
    )

    members_by_resource = {
        _access_id_key(resource.resource_id): []
        for resource in resources
        if resource.resource_id is not None
    }

    for emp in employees:
        resource_key = _access_id_key(emp.resource_id)

        if not resource_key:
            continue

        if resource_key not in members_by_resource:
            continue

        members_by_resource[resource_key].append(
            str(emp.employee_id)
        )

    conn = get_connection()
    cur = conn.cursor()

    try:
        for resource in resources:
            resource_key = _access_id_key(resource.resource_id)

            member_text = ", ".join(
                members_by_resource.get(resource_key, [])
            )

            cur.execute(
                """
                UPDATE resources
                SET resource_members = ?
                WHERE resource_id = ?
                """,
                member_text,
                resource.resource_id,
            )

        conn.commit()

    finally:
        conn.close()
########################################################################################################################
########################################################################################################################
########################################################################################################################





########################################################################################################################
########################################################################################################################
########################################################################################################################
"""LOGIN / SECURITY ROUTES"""
########################################################################################################################
########################################################################################################################
########################################################################################################################
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email_address = request.form.get("email_address", "").strip()
        password = request.form.get("password", "").strip()

        employee = fetch_one(
            """
            SELECT employee_id, employee_first_name, employee_last_name,
                   resource_id, email_address, [password], [office]
            FROM employees
            WHERE email_address = ?
            """,
            [email_address],
        )

        if not employee or str(employee.password) != password:
            flash("Invalid email or password.", "error")
            return redirect(url_for("login"))

        session["employee_id"] = employee.employee_id
        session["employee_name"] = f"{employee.employee_first_name} {employee.employee_last_name}"
        session["resource_id"] = employee.resource_id
        session["office"] = is_office_employee(employee)

        if is_office_employee(employee):
            return redirect(url_for("home"))

        open_day_punch = get_open_day_punch_for_employee(employee.employee_id)

        if not open_day_punch:
            return redirect(url_for("employee_day_clock_in"))

        open_punch = get_open_punch_for_employee(employee.employee_id)

        if open_punch:
            return redirect(url_for("employee_task_detail", task_id=open_punch.task_id))

        return redirect(url_for("employee_tasks"))

    html = """
    <!doctype html>
    <html>
    <head>
        <title>Signs App Login</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 40px; background: #f5f6f8; }
            .card { max-width: 440px; margin: 90px auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,.08); }
            input { width: 100%; padding: 10px; margin-top: 6px; margin-bottom: 14px; box-sizing: border-box; border: 1px solid #ccc; border-radius: 6px; }
            .btn { width: 100%; padding: 12px; background: #1f4e79; color: white; border: none; border-radius: 8px; font-weight: bold; cursor: pointer; }
            .flash-error { background: #f8d7da; color: #842029; padding: 10px; border-radius: 6px; margin-bottom: 12px; }
            h1 { margin-top: 0; }
        </style>
    </head>
    <body>
        <div class="card">
            <h1>Signs App</h1>
            <p>Log in with your employee email and password.</p>

            {% with messages = get_flashed_messages(with_categories=true) %}
                {% if messages %}
                    {% for category, message in messages %}
                        <div class="flash-{{ category }}">{{ message }}</div>
                    {% endfor %}
                {% endif %}
            {% endwith %}

            <form method="post">
                <label>Email Address</label>
                <input type="text" name="email_address" required>

                <label>Password</label>
                <input type="password" name="password" required>

                <button class="btn" type="submit">Login</button>
            </form>
        </div>
    </body>
    </html>
    """
    return render_template_string(html)
########################################################################################################################
@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))
########################################################################################################################
########################################################################################################################
########################################################################################################################




########################################################################################################################
########################################################################################################################
########################################################################################################################
"""HOME SCREEN"""
########################################################################################################################
########################################################################################################################
########################################################################################################################
@app.route("/")
@login_required
def home():
    employee = get_logged_in_employee()
    if not is_office_employee(employee):
        return redirect(url_for("employee_tasks"))

    gantt = get_home_gantt_data(weeks=6)
    task_table_rows = get_home_task_table_data()
    html = """
    <!doctype html>
    <html>
    <head>
        <title>Scheduling App</title>

<link rel="stylesheet"
      href="https://cdn.datatables.net/1.13.8/css/jquery.dataTables.min.css">

<script src="https://code.jquery.com/jquery-3.7.1.min.js"></script>

<script src="https://cdn.datatables.net/1.13.8/js/jquery.dataTables.min.js"></script>

<script>
    function toggleScheduleView() {
        const checked = document.getElementById("viewToggle").checked;
        document.getElementById("ganttView").style.display = checked ? "none" : "block";
        document.getElementById("taskTableView").style.display = checked ? "block" : "none";
    }

    $(document).ready(function () {

        const table = $('#taskTable').DataTable({
            pageLength: 25,
            orderCellsTop: true,
            fixedHeader: true
        });

        $('#taskTable thead tr:eq(1) th').each(function (i) {

            $('input', this).on('keyup change clear', function () {

                if (table.column(i).search() !== this.value) {
                    table.column(i).search(this.value).draw();
                }

            });

        });

    });
</script>
        <style>
            body {
                font-family: Arial, sans-serif;
                margin: 40px;
                background: #f5f6f8;
            }
            .card {
                background: white;
                padding: 30px;
                border-radius: 10px;
                box-shadow: 0 2px 10px rgba(0,0,0,.08);
                margin-bottom: 20px;
            }
            .btn {
                display: inline-block;
                padding: 12px 20px;
                background: #1f4e79;
                color: white;
                text-decoration: none;
                border-radius: 8px;
                font-weight: bold;
                margin-right: 10px;
                margin-bottom: 10px;
            }
            .view-toggle-wrap {
                display: flex;
                align-items: center;
                gap: 12px;
                margin-bottom: 18px;
            }
            .switch {
                position: relative;
                display: inline-block;
                width: 62px;
                height: 32px;
            }
            .switch input { display: none; }
            .slider {
                position: absolute;
                cursor: pointer;
                top: 0;
                left: 0;
                right: 0;
                bottom: 0;
                background-color: #1f4e79;
                transition: .2s;
                border-radius: 999px;
            }
            .slider:before {
                position: absolute;
                content: "";
                height: 24px;
                width: 24px;
                left: 4px;
                bottom: 4px;
                background-color: white;
                transition: .2s;
                border-radius: 50%;
            }
            input:checked + .slider:before {
                transform: translateX(30px);
            }
            .gantt-wrap {
    overflow-x: auto;
    overflow-y: visible;
    border: 1px solid #ddd;
    border-radius: 8px;
    background: #fff;
}
            .gantt-grid {
                min-width: calc(260px + {{ gantt.total_hour_columns }} * 46px);
            }
            .gantt-day-header, .gantt-hour-header, .gantt-row {
                display: grid;
                grid-template-columns: 260px repeat({{ gantt.total_hour_columns }}, 46px);
                align-items: stretch;
            }
            .gantt-cell {
                border-bottom: 1px solid #ddd;
                border-right: 1px solid #eee;
                padding: 4px;
                box-sizing: border-box;
            }
            .gantt-header-sticky {
                position: sticky;
                top: 0;
                z-index: 3;
                background: #fff;
            }
            .resource-label {
                font-weight: bold;
                font-size: 13px;
            }
            .timeline-cell {
                min-height: 34px;
                background: #fff;
            }
            .lunch-cell {
                background: #f1f3f5;
            }
            .day-header {
                text-align: center;
                font-weight: bold;
                background: #fafafa;
            }
            .hour-header {
                text-align: center;
                font-size: 11px;
                background: #fcfcfc;
                padding: 2px;
            }
            .gantt-row {
                position: relative;
            }
            .gantt-bar {
    position: absolute;
    top: 6px;
    height: 22px;
    background: #1f4e79;
    border-radius: 6px;
    color: white;
    padding: 4px 8px;
    box-sizing: border-box;

    overflow: visible;

    white-space: nowrap;
    text-overflow: ellipsis;
    font-size: 11px;
    line-height: 14px;

    cursor: pointer;
}

.gantt-link {
    text-decoration: none;
    color: inherit;
}

.gantt-tooltip {
    display: none;

    position: absolute;
    top: 28px;
    left: 0;

    min-width: 350px;
    max-width: 500px;

    background: white;
    color: black;

    border: 1px solid #ccc;
    border-radius: 8px;

    padding: 12px;

    box-shadow: 0 4px 12px rgba(0,0,0,.2);

    z-index: 99999;

    white-space: normal;
    line-height: 1.4;
}

.gantt-bar:hover .gantt-tooltip {
    display: block;
}

.tooltip-title {
    font-weight: bold;
    font-size: 14px;
    color: #1f4e79;
    margin-bottom: 8px;
}

.tooltip-row {
    margin-bottom: 4px;
}

.tooltip-label {
    font-weight: bold;
}
            .day-divider {
                border-left: 3px solid black !important;
            }
            .task-table-wrap {
                display: none;
                overflow-x: auto;
                border: 1px solid #ddd;
                border-radius: 8px;
                background: white;
            }
            .task-table {
                width: 100%;
                border-collapse: collapse;
            }
            .task-table th, .task-table td {
                padding: 8px;
                border-bottom: 1px solid #ddd;
                text-align: left;
                font-size: 13px;
                vertical-align: top;
            }
            .task-table th {
                background: #fafafa;
                font-weight: bold;
                position: sticky;
                top: 0;
                z-index: 2;
            }
            .task-description-cell {
                max-width: 380px;
                white-space: normal;
            }
            .task-table thead input {
    width: 100%;
    box-sizing: border-box;
    padding: 4px;
    font-size: 11px;
}

.dataTables_wrapper {
    padding: 10px;
}

.dataTables_filter {
    margin-bottom: 10px;
}
            h1 { margin-top: 0; }
            .page-wrap { max-width: 1400px; }
        </style>
    </head>
    <body>
        <div class="page-wrap">
            <div class="card">
                <h1>Scheduling App</h1>
                <p>Select a starting point.</p>
                {% if is_admin %}
                <a class="btn" href="{{ url_for('manage_templates') }}">Manage Tasking Templates</a>
                <a class="btn" href="{{ url_for('schedule_search') }}">Schedule Manufacturing</a>
                <a class="btn" href="{{ url_for('manual_task_assignment') }}">Assign Manual Task</a>
                <a class="btn" href="{{ url_for('installation_calendar') }}">Schedule Installation</a>
                <a class="btn" href="{{ url_for('resource_management') }}">Resource Management</a>
                <a class="btn" href="{{ url_for('manufacturing_floor_display') }}">Floor Display</a>
                <a class="btn"href="{{ url_for('qc_survey_responses') }}">QC and Survey Responses</a>
                {% endif %}
                <a class="btn" href="{{ url_for('logout') }}">Logout</a>
            </div>

            <div class="card">
                <h2>Job / WO Progress Lookup</h2>
                <form action="{{ url_for('job_lookup') }}" method="get" style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
                    <input type="text" name="q" placeholder="Enter Job ID or WO ID (e.g. 101 or WO 45)"
                           style="padding:10px 14px; border:1px solid #ccc; border-radius:6px; font-size:14px; width:300px;">
                    <button type="submit" class="btn" style="padding:10px 20px;">Look Up</button>
                </form>
            </div>

            <div class="card">
                <h2>Current Schedule</h2>
                <div class="view-toggle-wrap">
                    <strong>Gantt View</strong>
                    <label class="switch">
                        <input type="checkbox" id="viewToggle" onchange="toggleScheduleView()">
                        <span class="slider"></span>
                    </label>
                    <strong>Task View</strong>
                </div>

                <div id="ganttView">
                    <p>Showing scheduled work from {{ gantt.week_start_label }} through {{ gantt.horizon_end_label }}. Scroll right to see further into the future.</p>
                    <div class="gantt-wrap">
                        <div class="gantt-grid">
                            <div class="gantt-day-header gantt-header-sticky">
                                <div class="gantt-cell day-header">Resource</div>
                                {% for day in gantt.day_columns %}
                                    <div class="gantt-cell day-header day-divider" style="grid-column: span 9;">{{ day.label_top }}<br>{{ day.label_bottom }}</div>
                                {% endfor %}
                            </div>
                            <div class="gantt-hour-header gantt-header-sticky">
                                <div class="gantt-cell hour-header"></div>
                                {% for day in gantt.day_columns %}
                                    {% for hour in day.hours %}
                                        <div class="gantt-cell hour-header {% if hour == '12p' %}lunch-cell{% endif %} {% if loop.index0 == 0 %}day-divider{% endif %}">{{ hour }}</div>
                                    {% endfor %}
                                {% endfor %}
                            </div>

                            {% for row in gantt.gantt_rows %}
                                <div class="gantt-row">
                                    <div class="gantt-cell">
                                        <div class="resource-label">{{ row.resource_name }}</div>
                                    </div>
                                    {% for day in gantt.day_columns %}
                                        {% for hour in day.hours %}
                                            <div class="gantt-cell timeline-cell {% if hour == '12p' %}lunch-cell{% endif %} {% if loop.index0 == 0 %}day-divider{% endif %}"></div>
                                        {% endfor %}
                                    {% endfor %}
                                    {% for task in row.tasks %}
                                        <a class="gantt-link" href="{{ url_for('office_task_detail', task_id=task.task_id) }}">
    <div class="gantt-bar"
         style="left: calc(260px + {{ task.start_pos * 46 }}px + 3px); width: calc({{ task.span_pos * 46 }}px - 6px);">

        {{ task.task_name }}

        <div class="gantt-tooltip">
            <div class="tooltip-title">{{ task.task_name }}</div>
            <div class="tooltip-row"><span class="tooltip-label">Task ID:</span> {{ task.task_id }}</div>
            <div class="tooltip-row"><span class="tooltip-label">Description:</span> {{ task.task_description }}</div>
            <div class="tooltip-row"><span class="tooltip-label">Planned Hours:</span> {{ task.planned_hours }}</div>
            <div class="tooltip-row"><span class="tooltip-label">Resource:</span> {{ row.resource_name }}</div>
            <div class="tooltip-row"><span class="tooltip-label">Start:</span> {{ task.start_label }}</div>
            <div class="tooltip-row"><span class="tooltip-label">Finish:</span> {{ task.finish_label }}</div>
        </div>
    </div>
</a>
                                    {% endfor %}
                                </div>
                            {% else %}
                                <div class="gantt-row">
                                    <div class="gantt-cell">No scheduled tasks yet.</div>
                                    {% for day in gantt.day_columns %}
                                        {% for hour in day.hours %}
                                            <div class="gantt-cell timeline-cell {% if hour == '12p' %}lunch-cell{% endif %} {% if loop.index0 == 0 %}day-divider{% endif %}"></div>
                                        {% endfor %}
                                    {% endfor %}
                                </div>
                            {% endfor %}
                        </div>
                    </div>
                </div>

 <div id="taskTableView" class="task-table-wrap">

    <table id="taskTable" class="task-table">

        <thead>

            <tr>
                <th>Job / WO</th>
                <th>Task ID</th>
                <th>Cost Code</th>
                <th>Task Name</th>
                <th>Description</th>
                <th>Planned Hours</th>
                <th>Current Hours</th>
                <th>% Complete</th>
                <th>Resource Type</th>
                <th>Resource</th>
                <th>Start</th>
                <th>Finish</th>
                <th>Status</th>
                <th>Project Manager</th>
            </tr>

            <tr>
                <th><input type="text" placeholder="Search"></th>
                <th><input type="text" placeholder="Search"></th>
                <th><input type="text" placeholder="Search"></th>
                <th><input type="text" placeholder="Search"></th>
                <th><input type="text" placeholder="Search"></th>
                <th><input type="text" placeholder="Search"></th>
                <th><input type="text" placeholder="Search"></th>
                <th><input type="text" placeholder="Search"></th>
                <th><input type="text" placeholder="Search"></th>
                <th><input type="text" placeholder="Search"></th>
                <th><input type="text" placeholder="Search"></th>
                <th><input type="text" placeholder="Search"></th>
                <th><input type="text" placeholder="Search"></th>
                <th><input type="text" placeholder="Search"></th>
            </tr>

        </thead>

        <tbody>

            {% for row in task_table_rows %}
            <tr>
                <td>
                    <a href="{{ url_for('job_wo_detail', task_id=row.task_id) }}">
                    {{ row.job_wo_number }}
                    </a>
                </td>
                <td>
                    <a href="{{ url_for('office_task_detail', task_id=row.task_id) }}">
                    {{ row.task_id }}
                    </a>
                </td>
                <td>{{ row.cost_code }}</td>
                <td>{{ row.task_name }}</td>
                <td class="task-description-cell">{{ row.task_description }}</td>
                <td>{{ "%.2f"|format(row.planned_hours) }}</td>
                <td>{{ "%.2f"|format(row.current_hours) }}</td>
                <td>{{ "%.1f"|format(row.percent_complete) }}%</td>
                <td>{{ row.resource_type }}</td>
                <td>{{ row.resource }}</td>
                <td>{{ row.planned_start }}</td>
                <td>{{ row.planned_finish }}</td>
                <td>{{ row.status }}</td>
                <td>{{ row.project_manager }}</td>
            </tr>
            {% else %}
            <tr>
                <td colspan="11">No tasks found.</td>
            </tr>
            {% endfor %}

        </tbody>

    </table>

</div>
            </div>
        </div>
    </body>
    </html>
    """
    return render_template_string(html, gantt=gantt, task_table_rows=task_table_rows, is_admin=is_admin_employee(employee))
########################################################################################################################
@app.route("/office/tasks/<int:task_id>", methods=["GET", "POST"])
@login_required
@office_required
def office_task_detail(task_id):

    from types import SimpleNamespace

    employee = get_logged_in_employee()

    if not employee:
        return redirect(url_for("login"))

    # =========================================================================
    # LOAD TASK
    # =========================================================================

    task_row = fetch_one(
        """
        SELECT
            task_id,
            task_name,
            task_description,
            planned_hours,
            resource_type_id,
            resource_id,
            planned_start_date,
            planned_finish_date,
            status_id,
            percentage_complete,
            cost_code_id,
            survey_type
        FROM tasks
        WHERE task_id = ?
        """,
        [task_id],
    )

    if not task_row:
        flash("Task not found.", "error")
        return redirect(url_for("home"))

    # =========================================================================
    # LOAD COST CODE
    # =========================================================================

    cost_code = None

    if task_row.cost_code_id:

        cost_code = fetch_one(
            """
            SELECT
                cost_code_id,
                coins_cost_code,
                cost_code_description,
                line_item_id
            FROM cost_codes
            WHERE cost_code_id = ?
            """,
            [task_row.cost_code_id],
        )

    # =========================================================================
    # LOAD LINE ITEM / JOB / WO
    # =========================================================================

    line_item = None

    if cost_code and cost_code.line_item_id:

        line_item = fetch_one(
            """
            SELECT
                line_item_id,
                line_item_description,
                job_number_id,
                work_order_id
            FROM line_items
            WHERE line_item_id = ?
            """,
            [cost_code.line_item_id],
        )

    # =========================================================================
    # LOAD RESOURCE
    # =========================================================================

    assigned_resource = None

    if task_row.resource_id:

        assigned_resource = fetch_one(
            """
            SELECT
                resource_id,
                resource_name
            FROM resources
            WHERE resource_id = ?
            """,
            [task_row.resource_id],
        )

    # =========================================================================
    # LOAD RESOURCE TYPE
    # =========================================================================

    assigned_resource_type = None

    if task_row.resource_type_id:

        assigned_resource_type = fetch_one(
            """
            SELECT
                resource_type_id,
                resource_description
            FROM resource_types
            WHERE resource_type_id = ?
            """,
            [task_row.resource_type_id],
        )

    # =========================================================================
    # LOAD STATUS
    # =========================================================================

    status_row = None

    if task_row.status_id is not None:

        status_row = fetch_one(
            """
            SELECT
                status_id,
                status_names
            FROM task_status
            WHERE status_id = ?
            """,
            [task_row.status_id],
        )

    # =========================================================================
    # LOAD SURVEY / QC FORM NAME
    # =========================================================================

    survey_form = None

    if task_row.survey_type:

        survey_form = fetch_one(
            """
            SELECT
                qc_form_id,
                form_name
            FROM qc_forms
            WHERE qc_form_id = ?
            """,
            [task_row.survey_type],
        )

    # =========================================================================
    # COMBINE INTO TASK OBJECT
    # =========================================================================

    task = SimpleNamespace(

        task_id=task_row.task_id,
        task_name=task_row.task_name,
        task_description=task_row.task_description,
        planned_hours=task_row.planned_hours,

        resource_type_id=task_row.resource_type_id,
        resource_id=task_row.resource_id,

        planned_start_date=task_row.planned_start_date,
        planned_finish_date=task_row.planned_finish_date,

        status_id=task_row.status_id,

        percentage_complete=task_row.percentage_complete,

        cost_code_id=task_row.cost_code_id,

        survey_type=task_row.survey_type,

        coins_cost_code=(
            cost_code.coins_cost_code
            if cost_code
            else ""
        ),

        cost_code_description=(
            cost_code.cost_code_description
            if cost_code
            else ""
        ),

        line_item_id=(
            line_item.line_item_id
            if line_item
            else None
        ),

        line_item_description=(
            line_item.line_item_description
            if line_item
            else ""
        ),

        job_number_id=(
            line_item.job_number_id
            if line_item
            else None
        ),

        work_order_id=(
            line_item.work_order_id
            if line_item
            else None
        ),

        resource_name=(
            assigned_resource.resource_name
            if assigned_resource
            else ""
        ),

        resource_type_name=(
            assigned_resource_type.resource_description
            if assigned_resource_type
            else ""
        ),

        status_name=(
            status_row.status_names
            if status_row
            else ""
        ),

        survey_form_name=(
            survey_form.form_name
            if survey_form
            else ""
        ),
    )

    # =========================================================================
    # SAVE TASK EDITS
    # =========================================================================

    if request.method == "POST":

        task_name = request.form.get(
            "task_name",
            "",
        ).strip()

        task_description = request.form.get(
            "task_description",
            "",
        ).strip()

        planned_hours = request.form.get(
            "planned_hours",
            "",
        ).strip()

        resource_type_id = request.form.get(
            "resource_type_id",
            "",
        ).strip()

        resource_id = request.form.get(
            "resource_id",
            "",
        ).strip()

        planned_start_date = request.form.get(
            "planned_start_date",
            "",
        ).strip()

        planned_finish_date = request.form.get(
            "planned_finish_date",
            "",
        ).strip()

        status_id = request.form.get(
            "status_id",
            "",
        ).strip()

        override_overlap = (
            request.form.get("override_overlap")
            == "yes"
        )

        try:

            planned_start_dt = (
                datetime.strptime(
                    planned_start_date,
                    "%Y-%m-%dT%H:%M",
                )
                if planned_start_date
                else None
            )

            planned_finish_dt = (
                datetime.strptime(
                    planned_finish_date,
                    "%Y-%m-%dT%H:%M",
                )
                if planned_finish_date
                else None
            )

        except ValueError:

            flash(
                "Invalid planned start or finish date.",
                "error",
            )

            return redirect(
                url_for(
                    "office_task_detail",
                    task_id=task_id,
                )
            )

        new_resource_id = (
            int(resource_id)
            if resource_id
            else None
        )

        if (
            planned_start_dt
            and planned_finish_dt
            and planned_finish_dt <= planned_start_dt
        ):

            flash(
                "Planned finish must be after planned start.",
                "error",
            )

            return redirect(
                url_for(
                    "office_task_detail",
                    task_id=task_id,
                )
            )

        # =====================================================================
        # OVERLAP CHECK
        # =====================================================================

        overlap_tasks = []

        if (
            planned_start_dt
            and planned_finish_dt
            and new_resource_id
        ):

            overlap_tasks = fetch_all(
                """
                SELECT
                    task_id,
                    task_name,
                    planned_start_date,
                    planned_finish_date
                FROM tasks
                WHERE task_id <> ?
                  AND resource_id = ?
                  AND planned_start_date IS NOT NULL
                  AND planned_finish_date IS NOT NULL
                  AND planned_start_date < ?
                  AND planned_finish_date > ?
                ORDER BY planned_start_date
                """,
                [
                    task_id,
                    new_resource_id,
                    planned_finish_dt,
                    planned_start_dt,
                ],
            )

        if overlap_tasks and not override_overlap:

            html = """
            <!doctype html>
            <html>

            <head>

                <title>
                    Schedule Overlap Warning
                </title>

                <style>

                    body {
                        font-family: Arial, sans-serif;
                        margin: 40px;
                        background: #f5f6f8;
                    }

                    .wrap {
                        max-width: 1000px;
                        margin: 0 auto;
                    }

                    .card {
                        background: white;
                        padding: 28px;
                        border-radius: 10px;
                        box-shadow: 0 2px 10px rgba(0,0,0,.08);
                    }

                    .warning {
                        background: #fff3cd;
                        color: #664d03;
                        padding: 16px;
                        border-radius: 8px;
                        margin-bottom: 20px;
                    }

                    .btn {
                        padding: 12px 18px;
                        background: #1f4e79;
                        color: white;
                        border: none;
                        border-radius: 8px;
                        font-weight: bold;
                        cursor: pointer;
                        text-decoration: none;
                        display: inline-block;
                    }

                    .btn-danger {
                        background: #842029;
                    }

                    table {
                        width: 100%;
                        border-collapse: collapse;
                    }

                    th,
                    td {
                        padding: 10px;
                        border-bottom: 1px solid #ddd;
                        text-align: left;
                    }

                </style>

            </head>

            <body>

                <div class="wrap">

                    <div class="card">

                        <h1>
                            Schedule Overlap Warning
                        </h1>

                        <div class="warning">

                            The schedule change overlaps with
                            existing task(s) for this resource.

                        </div>

                        <table>

                            <thead>

                                <tr>

                                    <th>
                                        Task ID
                                    </th>

                                    <th>
                                        Task Name
                                    </th>

                                    <th>
                                        Start
                                    </th>

                                    <th>
                                        Finish
                                    </th>

                                </tr>

                            </thead>

                            <tbody>

                                {% for row in overlap_tasks %}

                                <tr>

                                    <td>
                                        {{ row.task_id }}
                                    </td>

                                    <td>
                                        {{ row.task_name or '' }}
                                    </td>

                                    <td>
                                        {{ row.planned_start_date }}
                                    </td>

                                    <td>
                                        {{ row.planned_finish_date }}
                                    </td>

                                </tr>

                                {% endfor %}

                            </tbody>

                        </table>

                        <br>

                        <form method="post">

                            <input
                                type="hidden"
                                name="override_overlap"
                                value="yes"
                            >

                            <input
                                type="hidden"
                                name="task_name"
                                value="{{ task_name }}"
                            >

                            <input
                                type="hidden"
                                name="task_description"
                                value="{{ task_description }}"
                            >

                            <input
                                type="hidden"
                                name="planned_hours"
                                value="{{ planned_hours }}"
                            >

                            <input
                                type="hidden"
                                name="resource_type_id"
                                value="{{ resource_type_id }}"
                            >

                            <input
                                type="hidden"
                                name="resource_id"
                                value="{{ resource_id }}"
                            >

                            <input
                                type="hidden"
                                name="planned_start_date"
                                value="{{ planned_start_date }}"
                            >

                            <input
                                type="hidden"
                                name="planned_finish_date"
                                value="{{ planned_finish_date }}"
                            >

                            <input
                                type="hidden"
                                name="status_id"
                                value="{{ status_id }}"
                            >

                            <button
                                class="btn btn-danger"
                                type="submit"
                            >
                                Override and Save Anyway
                            </button>

                            <a
                                class="btn"
                                href="{{
                                    url_for(
                                        'office_task_detail',
                                        task_id=task_id
                                    )
                                }}"
                            >
                                Cancel
                            </a>

                        </form>

                    </div>

                </div>

            </body>

            </html>
            """

            return render_template_string(

                html,

                task_id=task_id,

                overlap_tasks=overlap_tasks,

                task_name=task_name,

                task_description=task_description,

                planned_hours=planned_hours,

                resource_type_id=resource_type_id,

                resource_id=resource_id,

                planned_start_date=planned_start_date,

                planned_finish_date=planned_finish_date,

                status_id=status_id,
            )

        # =====================================================================
        # UPDATE TASK
        # =====================================================================

        conn = get_connection()
        cur = conn.cursor()

        cur.execute(
            """
            UPDATE tasks
            SET
                task_name = ?,
                task_description = ?,
                planned_hours = ?,
                resource_type_id = ?,
                resource_id = ?,
                planned_start_date = ?,
                planned_finish_date = ?,
                status_id = ?
            WHERE task_id = ?
            """,

            task_name,

            task_description,

            (
                float(planned_hours)
                if planned_hours
                else None
            ),

            (
                int(resource_type_id)
                if resource_type_id
                else None
            ),

            (
                int(resource_id)
                if resource_id
                else None
            ),

            planned_start_dt,

            planned_finish_dt,

            (
                int(status_id)
                if status_id
                else None
            ),

            task_id,
        )

        conn.commit()
        conn.close()

        flash(
            "Task updated successfully.",
            "success",
        )

        return redirect(
            url_for(
                "office_task_detail",
                task_id=task_id,
            )
        )

    # =========================================================================
    # DROPDOWNS
    # =========================================================================

    resources = get_resources()

    resource_types = get_resource_types()

    statuses = fetch_all(
        """
        SELECT
            status_id,
            status_names
        FROM task_status
        ORDER BY status_id
        """
    )

    # =========================================================================
    # TIME PUNCHES
    # =========================================================================

    raw_punches = fetch_all(
        """
        SELECT
            time_entry_ID,
            employee_id,
            punch_in,
            punch_out,
            task_time
        FROM time_punches
        WHERE task_id = ?
        ORDER BY punch_in DESC
        """,
        [task_id],
    )

    time_punches = []

    labor_by_employee = {}

    total_labor_minutes = 0.0

    total_labor_cost = 0.0

    now_value = datetime.now()

    for punch in raw_punches:

        punch_employee = None

        if punch.employee_id:

            punch_employee = fetch_one(
                """
                SELECT
                    employee_id,
                    employee_first_name,
                    employee_last_name,
                    burden_rate
                FROM employees
                WHERE employee_id = ?
                """,
                [punch.employee_id],
            )

        employee_name = (

            (
                f"{punch_employee.employee_first_name or ''} "
                f"{punch_employee.employee_last_name or ''}"
            ).strip()

            if punch_employee

            else f"Employee {punch.employee_id}"
        )

        burden_rate = (

            float(
                punch_employee.burden_rate or 0
            )

            if punch_employee

            else 0.0
        )

        if punch.task_time is not None:

            minutes = float(
                punch.task_time or 0
            )

        elif punch.punch_in:

            minutes = max(
                0,
                (
                    now_value - punch.punch_in
                ).total_seconds() / 60.0,
            )

        else:

            minutes = 0.0

        hours = minutes / 60.0

        labor_cost = (
            hours
            * burden_rate
        )

        total_labor_minutes += minutes

        total_labor_cost += labor_cost

        time_punches.append(

            SimpleNamespace(

                employee_name=employee_name,

                punch_in=punch.punch_in,

                punch_out=punch.punch_out,

                hours=hours,

                burden_rate=burden_rate,

                labor_cost=labor_cost,

                is_open=(
                    punch.punch_out is None
                ),
            )
        )

        summary = labor_by_employee.setdefault(

            punch.employee_id,

            {
                "employee_name": employee_name,
                "hours": 0.0,
                "labor_cost": 0.0,
                "punch_count": 0,
            },
        )

        summary["hours"] += hours

        summary["labor_cost"] += labor_cost

        summary["punch_count"] += 1

    labor_employee_rows = list(
        labor_by_employee.values()
    )

    total_labor_hours = (
        total_labor_minutes
        / 60.0
    )

    planned_hours_value = float(
        task.planned_hours or 0
    )

    labor_variance_hours = (
        total_labor_hours
        - planned_hours_value
    )

    # =========================================================================
    # MATERIALS
    #
    # This still uses cost_code_id because that is how your current
    # material_transactions data was historically stored.
    # =========================================================================

    raw_material_transactions = []

    if task.cost_code_id:

        raw_material_transactions = fetch_all(
            """
            SELECT
                material_id,
                quantity,
                transaction_date,
                employee_id
            FROM material_transactions
            WHERE cost_code_id = ?
            ORDER BY transaction_date DESC
            """,
            [task.cost_code_id],
        )

    material_transactions = []

    total_material_cost = 0.0

    for transaction in raw_material_transactions:

        material = None

        if transaction.material_id:

            material = fetch_one(
                """
                SELECT
                    material_id,
                    material_name,
                    barcode,
                    unit_cost
                FROM materials
                WHERE material_id = ?
                """,
                [transaction.material_id],
            )

        issued_by = ""

        if transaction.employee_id:

            issue_employee = fetch_one(
                """
                SELECT
                    employee_first_name,
                    employee_last_name
                FROM employees
                WHERE employee_id = ?
                """,
                [transaction.employee_id],
            )

            if issue_employee:

                issued_by = (

                    f"{issue_employee.employee_first_name or ''} "
                    f"{issue_employee.employee_last_name or ''}"

                ).strip()

        quantity = float(
            transaction.quantity or 0
        )

        unit_cost = (

            float(
                material.unit_cost or 0
            )

            if material

            else 0.0
        )

        extended_cost = (
            quantity
            * unit_cost
        )

        total_material_cost += (
            extended_cost
        )

        material_transactions.append(

            SimpleNamespace(

                material_id=(
                    transaction.material_id
                ),

                material_name=(
                    material.material_name
                    if material
                    else ""
                ),

                barcode=(
                    material.barcode
                    if material
                    else ""
                ),

                quantity=quantity,

                unit_cost=unit_cost,

                extended_cost=extended_cost,

                employee_name=issued_by,

                transaction_date=(
                    transaction.transaction_date
                ),
            )
        )

    # =========================================================================
    # QC / SURVEY ANSWERS
    #
    # Surveys now use qc_form_answers too.
    # =========================================================================

    raw_answers = fetch_all(
        """
        SELECT
            answer_value,
            submitted_date,
            qc_question_id,
            employee_id
        FROM qc_form_answers
        WHERE task_id = ?
        ORDER BY
            submitted_date,
            qc_question_id
        """,
        [task_id],
    )

    qc_answers = []

    for answer in raw_answers:

        question_text = ""

        if answer.qc_question_id:

            question = fetch_one(
                """
                SELECT question_text
                FROM qc_form_questions
                WHERE qc_question_id = ?
                """,
                [answer.qc_question_id],
            )

            if question:

                question_text = (
                    question.question_text
                    or ""
                )

        answer_employee_name = ""

        if answer.employee_id:

            answer_employee = fetch_one(
                """
                SELECT
                    employee_first_name,
                    employee_last_name
                FROM employees
                WHERE employee_id = ?
                """,
                [answer.employee_id],
            )

            if answer_employee:

                answer_employee_name = (

                    f"{answer_employee.employee_first_name or ''} "
                    f"{answer_employee.employee_last_name or ''}"

                ).strip()

        qc_answers.append(

            SimpleNamespace(

                answer_value=(
                    answer.answer_value
                ),

                submitted_date=(
                    answer.submitted_date
                ),

                qc_question_id=(
                    answer.qc_question_id
                ),

                question_text=(
                    question_text
                ),

                employee_name=(
                    answer_employee_name
                ),
            )
        )

    # =========================================================================
    # TOTALS
    # =========================================================================

    total_task_cost = (
        total_labor_cost
        + total_material_cost
    )

    percent_complete = float(
        task.percentage_complete or 0
    )

    # =========================================================================
    # DATE INPUT FORMAT
    # =========================================================================

    def fmt_dt_for_input(value):

        if not value:
            return ""

        return value.strftime(
            "%Y-%m-%dT%H:%M"
        )

    # =========================================================================
    # HTML
    # =========================================================================

    html = """
    <!doctype html>

    <html>

    <head>

        <title>
            Task {{ task.task_id }}
        </title>

        <style>

            * {
                box-sizing: border-box;
            }

            body {
                font-family: Arial, sans-serif;
                margin: 0;
                padding: 28px;
                background: #f5f6f8;
            }

            .wrap {
                max-width: 1400px;
                margin: auto;
            }

            .card {
                background: white;
                padding: 24px;
                border-radius: 10px;
                box-shadow: 0 2px 10px rgba(0,0,0,.08);
                margin-bottom: 20px;
            }

            .hero {
                background: #1f4e79;
                color: white;
            }

            .hero h1 {
                margin: 0 0 8px 0;
            }

            .btn {
                padding: 11px 18px;
                background: #1f4e79;
                color: white;
                border: none;
                border-radius: 7px;
                text-decoration: none;
                font-weight: bold;
                cursor: pointer;
                display: inline-block;
            }

            .gray {
                background: #6c757d;
            }

            .top-actions {
                display: flex;
                gap: 10px;
                flex-wrap: wrap;
                margin-bottom: 18px;
            }

            .summary-grid {
                display: grid;
                grid-template-columns:
                    repeat(
                        auto-fit,
                        minmax(175px,1fr)
                    );
                gap: 12px;
            }

            .metric {
                background: #f8f9fa;
                border: 1px solid #ddd;
                border-radius: 8px;
                padding: 16px;
            }

            .metric-label {
                color: #666;
                font-size: 12px;
                font-weight: bold;
                text-transform: uppercase;
            }

            .metric-value {
                font-size: 24px;
                color: #1f4e79;
                font-weight: bold;
                margin-top: 4px;
            }

            .info-grid {
                display: grid;
                grid-template-columns:
                    repeat(
                        auto-fit,
                        minmax(225px,1fr)
                    );
                gap: 12px;
            }

            .info {
                background: #fafafa;
                border: 1px solid #eee;
                padding: 12px;
                border-radius: 7px;
            }

            .info strong {
                display: block;
                color: #666;
                font-size: 12px;
                margin-bottom: 4px;
            }

            table {
                width: 100%;
                border-collapse: collapse;
            }

            th,
            td {
                padding: 9px;
                border-bottom: 1px solid #ddd;
                text-align: left;
                vertical-align: top;
            }

            th {
                background: #fafafa;
            }

            .number {
                text-align: right;
            }

            .grid {
                display: grid;
                grid-template-columns: 1fr 1fr;
                gap: 18px;
            }

            label {
                display: block;
                font-weight: bold;
                margin-bottom: 5px;
            }

            input,
            textarea,
            select {
                width: 100%;
                padding: 10px;
                border: 1px solid #ccc;
                border-radius: 6px;
                margin-bottom: 14px;
            }

            textarea {
                min-height: 90px;
            }

            .clocked-in {
                color: #198754;
                font-weight: bold;
            }

            .flash-success {
                background: #d1e7dd;
                color: #0f5132;
                padding: 10px;
                border-radius: 6px;
                margin-bottom: 12px;
            }

            .flash-error {
                background: #f8d7da;
                color: #842029;
                padding: 10px;
                border-radius: 6px;
                margin-bottom: 12px;
            }

            @media(max-width:800px) {

                body {
                    padding: 12px;
                }

                .grid {
                    grid-template-columns: 1fr;
                }

                .table-wrap {
                    overflow-x: auto;
                }
            }

        </style>

    </head>

    <body>

        <div class="wrap">

            <div class="top-actions">

                <a
                    class="btn gray"
                    href="{{ url_for('home') }}"
                >
                    ← Back to Schedule
                </a>

                <a
                    class="btn"
                    href="{{
                        url_for(
                            'task_files',
                            task_id=task.task_id
                        )
                    }}"
                >
                    Task Files
                </a>

            </div>

            {% with messages = get_flashed_messages(with_categories=true) %}

                {% if messages %}

                    {% for category, message in messages %}

                        <div class="flash-{{ category }}">
                            {{ message }}
                        </div>

                    {% endfor %}

                {% endif %}

            {% endwith %}


            <div class="card hero">

                <h1>
                    Task {{ task.task_id }}
                    — {{ task.task_name or 'Unnamed Task' }}
                </h1>

                {% if task.job_number_id %}

                    Job {{ task.job_number_id }}

                {% elif task.work_order_id %}

                    WO {{ task.work_order_id }}

                {% endif %}

                {% if task.line_item_description %}

                    —
                    {{ task.line_item_description }}

                {% endif %}

            </div>


            <div class="card">

                <h2>
                    Task Summary
                </h2>

                <div class="summary-grid">

                    <div class="metric">

                        <div class="metric-label">
                            Planned Hours
                        </div>

                        <div class="metric-value">
                            {{ "%.2f"|format(planned_hours_value) }}
                        </div>

                    </div>


                    <div class="metric">

                        <div class="metric-label">
                            Actual Hours
                        </div>

                        <div class="metric-value">
                            {{ "%.2f"|format(total_labor_hours) }}
                        </div>

                    </div>


                    <div class="metric">

                        <div class="metric-label">
                            Hour Variance
                        </div>

                        <div class="metric-value">
                            {{ "%+.2f"|format(labor_variance_hours) }}
                        </div>

                    </div>


                    <div class="metric">

                        <div class="metric-label">
                            Labor Cost
                        </div>

                        <div class="metric-value">
                            ${{ "{:,.2f}".format(total_labor_cost) }}
                        </div>

                    </div>


                    <div class="metric">

                        <div class="metric-label">
                            Material Cost
                        </div>

                        <div class="metric-value">
                            ${{ "{:,.2f}".format(total_material_cost) }}
                        </div>

                    </div>


                    <div class="metric">

                        <div class="metric-label">
                            Labor + Material
                        </div>

                        <div class="metric-value">
                            ${{ "{:,.2f}".format(total_task_cost) }}
                        </div>

                    </div>

                </div>

            </div>


            <div class="card">

                <h2>
                    Task Information
                </h2>

                <div class="info-grid">

                    <div class="info">

                        <strong>
                            Job / WO
                        </strong>

                        {% if task.job_number_id %}

                            Job {{ task.job_number_id }}

                        {% elif task.work_order_id %}

                            WO {{ task.work_order_id }}

                        {% else %}

                            —

                        {% endif %}

                    </div>


                    <div class="info">

                        <strong>
                            Cost Code
                        </strong>

                        {{
                            task.coins_cost_code
                            or task.cost_code_description
                            or '—'
                        }}

                    </div>


                    <div class="info">

                        <strong>
                            Resource Type
                        </strong>

                        {{ task.resource_type_name or '—' }}

                    </div>


                    <div class="info">

                        <strong>
                            Resource
                        </strong>

                        {{ task.resource_name or '—' }}

                    </div>


                    <div class="info">

                        <strong>
                            Status
                        </strong>

                        {{ task.status_name or task.status_id }}

                    </div>


                    <div class="info">

                        <strong>
                            Percent Complete
                        </strong>

                        {{ "%.1f"|format(percent_complete) }}%

                    </div>


                    <div class="info">

                        <strong>
                            Planned Start
                        </strong>

                        {{
                            task.planned_start_date.strftime(
                                '%b %d, %Y %I:%M %p'
                            )
                            if task.planned_start_date
                            else '—'
                        }}

                    </div>


                    <div class="info">

                        <strong>
                            Planned Finish
                        </strong>

                        {{
                            task.planned_finish_date.strftime(
                                '%b %d, %Y %I:%M %p'
                            )
                            if task.planned_finish_date
                            else '—'
                        }}

                    </div>


                    {% if task.survey_type %}

                    <div class="info">

                        <strong>
                            Survey Type
                        </strong>

                        {{
                            task.survey_form_name
                            or task.survey_type
                        }}

                    </div>

                    {% endif %}

                </div>

            </div>


            <div class="card">

                <h2>
                    Labor By Employee
                </h2>

                <div class="table-wrap">

                    <table>

                        <thead>

                            <tr>
                                <th>Employee</th>
                                <th>Punches</th>
                                <th>Hours</th>
                                <th>Labor Cost</th>
                            </tr>

                        </thead>

                        <tbody>

                            {% for row in labor_employee_rows %}

                            <tr>

                                <td>
                                    {{ row.employee_name }}
                                </td>

                                <td>
                                    {{ row.punch_count }}
                                </td>

                                <td>
                                    {{ "%.2f"|format(row.hours) }}
                                </td>

                                <td>
                                    ${{ "{:,.2f}".format(row.labor_cost) }}
                                </td>

                            </tr>

                            {% else %}

                            <tr>

                                <td colspan="4">
                                    No labor recorded.
                                </td>

                            </tr>

                            {% endfor %}

                        </tbody>

                    </table>

                </div>

            </div>


            <div class="card">

                <h2>
                    Time Punch History
                </h2>

                <div class="table-wrap">

                    <table>

                        <thead>

                            <tr>
                                <th>Employee</th>
                                <th>Punch In</th>
                                <th>Punch Out</th>
                                <th>Hours</th>
                                <th>Labor Cost</th>
                            </tr>

                        </thead>

                        <tbody>

                            {% for p in time_punches %}

                            <tr>

                                <td>
                                    {{ p.employee_name }}
                                </td>

                                <td>

                                    {{
                                        p.punch_in.strftime(
                                            '%b %d, %Y %I:%M %p'
                                        )
                                        if p.punch_in
                                        else '—'
                                    }}

                                </td>

                                <td>

                                    {% if p.is_open %}

                                        <span class="clocked-in">
                                            CLOCKED IN
                                        </span>

                                    {% else %}

                                        {{
                                            p.punch_out.strftime(
                                                '%b %d, %Y %I:%M %p'
                                            )
                                            if p.punch_out
                                            else '—'
                                        }}

                                    {% endif %}

                                </td>

                                <td>
                                    {{ "%.2f"|format(p.hours) }}
                                </td>

                                <td>
                                    ${{ "{:,.2f}".format(p.labor_cost) }}
                                </td>

                            </tr>

                            {% else %}

                            <tr>

                                <td colspan="5">
                                    No time punches found.
                                </td>

                            </tr>

                            {% endfor %}

                        </tbody>

                    </table>

                </div>

            </div>


            <div class="card">

                <h2>
                    Materials Costed
                </h2>

                <div class="table-wrap">

                    <table>

                        <thead>

                            <tr>
                                <th>Material</th>
                                <th>Barcode</th>
                                <th>Quantity</th>
                                <th>Unit Cost</th>
                                <th>Total Cost</th>
                                <th>Issued By</th>
                                <th>Date</th>
                            </tr>

                        </thead>

                        <tbody>

                            {% for m in material_transactions %}

                            <tr>

                                <td>
                                    {{
                                        m.material_name
                                        or (
                                            'Material '
                                            ~ m.material_id
                                        )
                                    }}
                                </td>

                                <td>
                                    {{ m.barcode or '—' }}
                                </td>

                                <td>
                                    {{ "%.2f"|format(m.quantity) }}
                                </td>

                                <td>
                                    ${{ "{:,.2f}".format(m.unit_cost) }}
                                </td>

                                <td>
                                    ${{ "{:,.2f}".format(m.extended_cost) }}
                                </td>

                                <td>
                                    {{ m.employee_name or '—' }}
                                </td>

                                <td>

                                    {{
                                        m.transaction_date.strftime(
                                            '%b %d, %Y %I:%M %p'
                                        )
                                        if m.transaction_date
                                        else '—'
                                    }}

                                </td>

                            </tr>

                            {% else %}

                            <tr>

                                <td colspan="7">
                                    No material transactions found.
                                </td>

                            </tr>

                            {% endfor %}

                        </tbody>

                    </table>

                </div>

            </div>


            {% if qc_answers %}

            <div class="card">

                <h2>
                    QC / Survey Answers
                </h2>

                {% set ns = namespace(last_date=None) %}

                {% for a in qc_answers %}

                    {% if a.submitted_date != ns.last_date %}

                        {% if not loop.first %}
                            </div>
                        {% endif %}

                        <div
                            style="
                                border:1px solid #ddd;
                                padding:16px;
                                border-radius:8px;
                                margin-bottom:15px;
                            "
                        >

                            <div
                                style="
                                    color:#666;
                                    font-size:13px;
                                    margin-bottom:10px;
                                "
                            >

                                Submitted by

                                <strong>
                                    {{ a.employee_name or 'Unknown' }}
                                </strong>

                                on

                                {{
                                    a.submitted_date.strftime(
                                        '%b %d, %Y %I:%M %p'
                                    )
                                    if a.submitted_date
                                    else '—'
                                }}

                            </div>

                        {% set ns.last_date =
                            a.submitted_date %}

                    {% endif %}


                    {% if
                        a.answer_value
                        and a.answer_value.startswith(
                            'NOTES:'
                        )
                    %}

                        <p>

                            <strong>
                                Notes:
                            </strong>

                            {{ a.answer_value[6:] }}

                        </p>

                    {% elif a.question_text %}

                        <p>

                            <strong>
                                {{ a.question_text }}:
                            </strong>

                            {{ a.answer_value or '—' }}

                        </p>

                    {% endif %}

                {% endfor %}

                {% if qc_answers %}
                    </div>
                {% endif %}

            </div>

            {% endif %}


            <div class="card">

                <h2>
                    Edit Task
                </h2>

                <form method="post">

                    <div class="grid">

                        <div>

                            <label>
                                Task Name
                            </label>

                            <input
                                type="text"
                                name="task_name"
                                value="{{ task.task_name or '' }}"
                            >


                            <label>
                                Task Description
                            </label>

                            <textarea
                                name="task_description"
                            >{{ task.task_description or '' }}</textarea>


                            <label>
                                Planned Hours
                            </label>

                            <input
                                type="number"
                                step="0.01"
                                name="planned_hours"
                                value="{{ task.planned_hours or '' }}"
                            >


                            <label>
                                Status
                            </label>

                            <select name="status_id">

                                {% for s in statuses %}

                                    <option
                                        value="{{ s.status_id }}"
                                        {% if s.status_id == task.status_id %}
                                            selected
                                        {% endif %}
                                    >
                                        {{ s.status_names }}
                                    </option>

                                {% endfor %}

                            </select>

                        </div>


                        <div>

                            <label>
                                Resource Type
                            </label>

                            <select name="resource_type_id">

                                <option value="">
                                    -- None --
                                </option>

                                {% for rt in resource_types %}

                                    <option
                                        value="{{ rt.resource_type_id }}"
                                        {% if rt.resource_type_id == task.resource_type_id %}
                                            selected
                                        {% endif %}
                                    >
                                        {{ rt.resource_description }}
                                    </option>

                                {% endfor %}

                            </select>


                            <label>
                                Resource
                            </label>

                            <select name="resource_id">

                                <option value="">
                                    -- None --
                                </option>

                                {% for r in resources %}

                                    <option
                                        value="{{ r.resource_id }}"
                                        {% if r.resource_id == task.resource_id %}
                                            selected
                                        {% endif %}
                                    >
                                        {{ r.resource_name }}
                                    </option>

                                {% endfor %}

                            </select>


                            <label>
                                Planned Start
                            </label>

                            <input
                                type="datetime-local"
                                name="planned_start_date"
                                value="{{ planned_start_value }}"
                            >


                            <label>
                                Planned Finish
                            </label>

                            <input
                                type="datetime-local"
                                name="planned_finish_date"
                                value="{{ planned_finish_value }}"
                            >

                        </div>

                    </div>


                    <button
                        class="btn"
                        type="submit"
                    >
                        Save Task
                    </button>

                </form>

            </div>

        </div>

    </body>

    </html>
    """

    return render_template_string(

        html,

        task=task,

        resources=resources,

        resource_types=resource_types,

        statuses=statuses,

        time_punches=time_punches,

        labor_employee_rows=(
            labor_employee_rows
        ),

        material_transactions=(
            material_transactions
        ),

        qc_answers=qc_answers,

        planned_start_value=(
            fmt_dt_for_input(
                task.planned_start_date
            )
        ),

        planned_finish_value=(
            fmt_dt_for_input(
                task.planned_finish_date
            )
        ),

        planned_hours_value=(
            planned_hours_value
        ),

        total_labor_hours=(
            total_labor_hours
        ),

        labor_variance_hours=(
            labor_variance_hours
        ),

        total_labor_cost=(
            total_labor_cost
        ),

        total_material_cost=(
            total_material_cost
        ),

        total_task_cost=(
            total_task_cost
        ),

        percent_complete=(
            percent_complete
        ),
    )
########################################################################################################################
@app.route("/job-lookup")
@login_required
@office_required
def job_lookup():
    q = request.args.get("q", "").strip()
    if not q:
        return redirect(url_for("home"))

    raw = q.upper().replace("WO", "").replace("JOB", "").strip()
    try:
        entity_id = int(raw)
    except ValueError:
        flash(f"'{q}' is not a valid Job or WO ID.", "error")
        return redirect(url_for("home"))

    job = fetch_one("SELECT job_id, due_date, customer_id FROM jobs WHERE job_id = ?", [entity_id])
    wo  = fetch_one("SELECT wo_id,  due_date, customer_id FROM work_orders WHERE wo_id = ?",  [entity_id])

    results = []

    if job:
        customer = fetch_one("SELECT customer_name FROM customers WHERE customer_id = ?", [job.customer_id])
        mfg_tasks = fetch_all(
            """
            SELECT t.task_id, t.task_name, t.planned_hours, t.planned_start_date, t.planned_finish_date,
                   t.status_id, r.resource_name, cc.cost_code_description
            FROM (((line_items AS li
                INNER JOIN cost_codes AS cc ON cc.line_item_id = li.line_item_id)
                INNER JOIN tasks AS t ON t.cost_code_id = cc.cost_code_id)
                LEFT JOIN resources AS r ON t.resource_id = r.resource_id)
            WHERE li.job_number_id = ?
              AND t.resource_type_id <> ?
            ORDER BY t.planned_start_date, t.task_id
            """,
            [entity_id, INSTALLATION_RESOURCE_TYPE_ID],
        )
        install_task = fetch_one(
            """
            SELECT TOP 1 t.planned_start_date, t.planned_finish_date, t.status_id, r.resource_name
            FROM (((line_items AS li
                INNER JOIN cost_codes AS cc ON cc.line_item_id = li.line_item_id)
                INNER JOIN tasks AS t ON t.cost_code_id = cc.cost_code_id)
                LEFT JOIN resources AS r ON t.resource_id = r.resource_id)
            WHERE li.job_number_id = ?
              AND t.resource_type_id = ?
            ORDER BY t.task_id
            """,
            [entity_id, INSTALLATION_RESOURCE_TYPE_ID],
        )
        results.append({
            "label": f"Job {entity_id}",
            "entity_type": "job",
            "customer": customer.customer_name if customer else "—",
            "due_date": job.due_date.strftime("%B %d, %Y") if job.due_date else "—",
            "mfg_tasks": mfg_tasks,
            "install_task": install_task,
        })

    if wo:
        customer = fetch_one("SELECT customer_name FROM customers WHERE customer_id = ?", [wo.customer_id])
        mfg_tasks = fetch_all(
            """
            SELECT t.task_id, t.task_name, t.planned_hours, t.planned_start_date, t.planned_finish_date,
                   t.status_id, r.resource_name, cc.cost_code_description
            FROM (((line_items AS li
                INNER JOIN cost_codes AS cc ON cc.line_item_id = li.line_item_id)
                INNER JOIN tasks AS t ON t.cost_code_id = cc.cost_code_id)
                LEFT JOIN resources AS r ON t.resource_id = r.resource_id)
            WHERE li.work_order_id = ?
              AND t.resource_type_id <> ?
            ORDER BY t.planned_start_date, t.task_id
            """,
            [entity_id, INSTALLATION_RESOURCE_TYPE_ID],
        )
        install_task = fetch_one(
            """
            SELECT TOP 1 t.planned_start_date, t.planned_finish_date, t.status_id, r.resource_name
            FROM (((line_items AS li
                INNER JOIN cost_codes AS cc ON cc.line_item_id = li.line_item_id)
                INNER JOIN tasks AS t ON t.cost_code_id = cc.cost_code_id)
                LEFT JOIN resources AS r ON t.resource_id = r.resource_id)
            WHERE li.work_order_id = ?
              AND t.resource_type_id = ?
            ORDER BY t.task_id
            """,
            [entity_id, INSTALLATION_RESOURCE_TYPE_ID],
        )
        results.append({
            "label": f"WO {entity_id}",
            "entity_type": "wo",
            "customer": customer.customer_name if customer else "—",
            "due_date": wo.due_date.strftime("%B %d, %Y") if wo.due_date else "—",
            "mfg_tasks": mfg_tasks,
            "install_task": install_task,
        })

    if not results:
        flash(f"No Job or Work Order found with ID {entity_id}.", "error")
        return redirect(url_for("home"))

    def status_label(status_id):
        v = int(status_id or 0)
        if v == 3: return ("Complete",    "#198754")
        if v == 2: return ("In Progress", "#e07b00")
        return              ("Not Started", "#6c757d")

    html = """
    <!doctype html><html>
    <head>
        <title>Progress Lookup — {{ query }}</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 40px; background: #f5f6f8; }
            .wrap { max-width: 1100px; margin: 0 auto; }
            .card { background: white; padding: 24px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,.08); margin-bottom: 22px; }
            h1 { margin: 0 0 4px 0; color: #1f4e79; }
            h2 { margin: 0 0 14px 0; font-size: 16px; color: #333; border-bottom: 2px solid #e9ecef; padding-bottom: 6px; }
            .btn-back { padding: 10px 16px; background: #6c757d; color: white; text-decoration: none; border-radius: 8px; font-weight: bold; display: inline-block; margin-bottom: 18px; }
            table { width: 100%; border-collapse: collapse; font-size: 13px; }
            th, td { padding: 8px 10px; border-bottom: 1px solid #e9ecef; text-align: left; }
            th { background: #f8f9fa; font-weight: bold; color: #555; }
            .badge { display: inline-block; padding: 3px 9px; border-radius: 999px; font-size: 12px; font-weight: bold; color: white; }
            .meta { color: #666; font-size: 14px; margin: 2px 0; }
            .section-label { font-size: 13px; font-weight: bold; color: #1f4e79; margin: 16px 0 8px 0; }
            .install-row { display: flex; gap: 30px; flex-wrap: wrap; margin-top: 8px; font-size: 14px; }
            .install-field { }
            .install-field strong { display: block; font-size: 12px; color: #888; margin-bottom: 2px; }
        </style>
    </head>
    <body>
    <div class="wrap">
        <a class="btn-back" href="{{ url_for('home') }}">← Back</a>
        <h1 style="color:#1f4e79; margin-bottom:20px;">Progress: "{{ query }}"</h1>

        {% for r in results %}
        <div class="card">
            <h1>{{ r.label }}</h1>
            <p class="meta"><strong>Customer:</strong> {{ r.customer }}</p>
            <p class="meta"><strong>Due Date:</strong> {{ r.due_date }}</p>

            <div class="section-label">Manufacturing Tasks</div>
            {% if r.mfg_tasks %}
            <table>
                <thead>
                    <tr>
                        <th>Cost Code</th>
                        <th>Task</th>
                        <th>Resource</th>
                        <th>Planned Start</th>
                        <th>Planned Finish</th>
                        <th>Hours</th>
                        <th>Status</th>
                    </tr>
                </thead>
                <tbody>
                    {% for t in r.mfg_tasks %}
                    {% set sl = status_label(t.status_id) %}
                    <tr>
                        <td>{{ t.cost_code_description or '—' }}</td>
                        <td>{{ t.task_name or '—' }}</td>
                        <td>{{ t.resource_name or '—' }}</td>
                        <td>{{ t.planned_start_date.strftime('%b %d, %Y') if t.planned_start_date else '—' }}</td>
                        <td>{{ t.planned_finish_date.strftime('%b %d, %Y') if t.planned_finish_date else '—' }}</td>
                        <td>{{ t.planned_hours or '—' }}</td>
                        <td><span class="badge" style="background:{{ sl[1] }};">{{ sl[0] }}</span></td>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
            {% else %}
            <p style="color:#aaa; font-size:13px;">No manufacturing tasks scheduled.</p>
            {% endif %}

            <div class="section-label">Installation</div>
            {% if r.install_task and r.install_task.planned_start_date %}
            {% set sl = status_label(r.install_task.status_id) %}
            <div class="install-row">
                <div class="install-field"><strong>Crew</strong>{{ r.install_task.resource_name or '—' }}</div>
                <div class="install-field"><strong>Date</strong>{{ r.install_task.planned_start_date.strftime('%B %d, %Y') }}</div>
                <div class="install-field"><strong>Start</strong>{{ r.install_task.planned_start_date.strftime('%I:%M %p') }}</div>
                <div class="install-field"><strong>Finish</strong>{{ r.install_task.planned_finish_date.strftime('%I:%M %p') if r.install_task.planned_finish_date else '—' }}</div>
                <div class="install-field"><strong>Status</strong><span class="badge" style="background:{{ sl[1] }};">{{ sl[0] }}</span></div>
            </div>
            {% else %}
            <p style="color:#aaa; font-size:13px;">Not yet scheduled.</p>
            {% endif %}
        </div>
        {% endfor %}
    </div>
    </body></html>
    """

    return render_template_string(html, results=results, query=q, status_label=status_label)
########################################################################################################################
########################################################################################################################
########################################################################################################################





########################################################################################################################
########################################################################################################################
########################################################################################################################
"""TECH TASKING AND INVENTORY MANAGEMENT"""
########################################################################################################################
########################################################################################################################
########################################################################################################################
@app.route("/employee/day-clock-in", methods=["GET", "POST"])
@login_required
def employee_day_clock_in():
    employee = get_logged_in_employee()

    if not employee:
        return redirect(url_for("login"))

    if request.method == "POST":
        open_day_punch = get_open_day_punch_for_employee(employee.employee_id)

        if not open_day_punch:
            conn = get_connection()
            cur = conn.cursor()

            cur.execute(
                """
                INSERT INTO time_punches
                    (task_id, employee_id, punch_in, punch_out, task_time)
                VALUES (?, ?, ?, ?, ?)
                """,
                None,
                employee.employee_id,
                datetime.now(),
                None,
                0,
            )

            conn.commit()
            conn.close()

        return redirect(url_for("employee_tasks"))

    html = """
    <!doctype html>
    <html>
    <head>
        <title>Clock In</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 40px; background: #f5f6f8; }
            .wrap { max-width: 700px; margin: 0 auto; }
            .card { background: white; padding: 40px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,.08); text-align: center; }
            .btn { width: 100%; padding: 32px; background: #198754; color: white; border: none; border-radius: 8px; font-size: 28px; font-weight: bold; cursor: pointer; }
        </style>
    </head>
    <body>
        <div class="wrap">
            <div class="card">
                <h1>Start Day</h1>
                <form method="post">
                    <button class="btn" type="submit">Clock In For Day</button>
                </form>
            </div>
        </div>
    </body>
    </html>
    """

    return render_template_string(html)
########################################################################################################################
@app.route("/employee/day-clock-out", methods=["POST"])
@login_required
def employee_day_clock_out():
    employee = get_logged_in_employee()

    if not employee:
        return redirect(url_for("login"))

    now_value = datetime.now()

    open_task_punch = get_open_punch_for_employee(employee.employee_id)
    open_day_punch = get_open_day_punch_for_employee(employee.employee_id)

    conn = get_connection()
    cur = conn.cursor()

    if open_task_punch:
        task_minutes = int((now_value - open_task_punch.punch_in).total_seconds() / 60)

        cur.execute(
            """
            UPDATE time_punches
            SET punch_out = ?,
                task_time = ?
            WHERE time_entry_ID = ?
            """,
            now_value,
            task_minutes,
            open_task_punch.time_entry_ID,
        )

        update_task_id = open_task_punch.task_id
    else:
        update_task_id = None

    if open_day_punch:
        day_minutes = int((now_value - open_day_punch.punch_in).total_seconds() / 60)

        cur.execute(
            """
            UPDATE time_punches
            SET punch_out = ?,
                task_time = ?
            WHERE time_entry_ID = ?
            """,
            now_value,
            day_minutes,
            open_day_punch.time_entry_ID,
        )

    conn.commit()
    conn.close()

    if update_task_id:
        update_task_percentage_complete(update_task_id)

    flash("Clocked out for the day.", "success")
    return redirect(url_for("login"))
########################################################################################################################
@app.route("/my-tasks/<int:task_id>/clock-in", methods=["POST"])
@login_required
def clock_in_task(task_id):
    employee = get_logged_in_employee()

    if not employee:
        return redirect(url_for("login"))

    already_clocked_in = get_open_punch_for_employee(employee.employee_id)

    if already_clocked_in:
        return redirect(url_for("employee_task_detail", task_id=already_clocked_in.task_id))

    task = fetch_one(
        """
        SELECT task_id, resource_id, actual_start_date
        FROM tasks
        WHERE task_id = ?
          AND (
            resource_id = ?
            OR task_id IN (?, ?)
            OR resource_type_id = (SELECT resource_types FROM resources WHERE resource_id = ?)
          )
        """,
        [task_id, employee.resource_id, TECH_OPS_TASK_ID, OVERHEAD_TASK_ID, employee.resource_id],
    )

    if not task:
        flash("Task not found or not assigned to you.", "error")
        return redirect(url_for("employee_tasks"))

    in_progress_status_id = get_status_id_by_name("in_progress")

    if not in_progress_status_id:
        in_progress_status_id = get_status_id_by_name("In Progress")

    if not in_progress_status_id:
        flash("Could not find an in_progress status in task_status.", "error")
        return redirect(url_for("employee_tasks"))

    now_value = datetime.now()

    conn = get_connection()
    cur = conn.cursor()

    cur.execute(
        """
        INSERT INTO time_punches
            (task_id, employee_id, punch_in, punch_out, task_time)
        VALUES (?, ?, ?, ?, ?)
        """,
        task_id,
        employee.employee_id,
        now_value,
        None,
        0,
    )

    if task.actual_start_date is None:
        cur.execute(
            """
            UPDATE tasks
            SET status_id = ?,
                actual_start_date = ?
            WHERE task_id = ?
            """,
            in_progress_status_id,
            now_value,
            task_id,
        )
    else:
        cur.execute(
            """
            UPDATE tasks
            SET status_id = ?
            WHERE task_id = ?
            """,
            in_progress_status_id,
            task_id,
        )

    conn.commit()
    conn.close()

    if task_id in INDIRECT_TASK_IDS:
        return redirect(url_for("indirect_task_detail", task_id=task_id))

    return redirect(url_for("employee_task_detail", task_id=task_id))
########################################################################################################################
@app.route("/my-tasks/<int:task_id>/clock-out", methods=["POST"])
@login_required
def clock_out_task(task_id):
    employee = get_logged_in_employee()

    if not employee:
        return redirect(url_for("login"))

    open_punch = fetch_one(
        """
        SELECT TOP 1 time_entry_ID, punch_in
        FROM time_punches
        WHERE task_id = ?
          AND employee_id = ?
          AND punch_out IS NULL
        ORDER BY punch_in DESC
        """,
        [task_id, employee.employee_id],
    )

    if not open_punch:
        flash("No open punch found for this task.", "error")
        return redirect(url_for("employee_tasks"))

    now_value = datetime.now()
    minutes_worked = (now_value - open_punch.punch_in).total_seconds() / 60

    conn = get_connection()
    cur = conn.cursor()

    cur.execute(
        """
        UPDATE time_punches
        SET punch_out = ?,
            task_time = ?
        WHERE time_entry_ID = ?
        """,
        now_value,
        minutes_worked,
        open_punch.time_entry_ID,
    )

    conn.commit()
    conn.close()

    update_task_percentage_complete(task_id)

    flash("Clocked out successfully.", "success")
    return redirect(url_for("employee_tasks"))
########################################################################################################################
@app.route("/my-tasks/<int:task_id>/complete", methods=["POST"])
@login_required
def complete_task(task_id):
    employee = get_logged_in_employee()

    if not employee:
        return redirect(url_for("login"))

    task = fetch_one(
        """
        SELECT task_id, resource_id, resource_type_id
        FROM tasks
        WHERE task_id = ?
          AND (
              resource_id = ?
              OR resource_type_id = (
                  SELECT resource_types
                  FROM resources
                  WHERE resource_id = ?
              )
          )
        """,
        [task_id, employee.resource_id, employee.resource_id],
    )

    if not task:
        flash("Task not found or not assigned to you.", "error")
        return redirect(url_for("employee_tasks"))

    complete_status_id = get_status_id_by_name("complete")

    if not complete_status_id:
        complete_status_id = get_status_id_by_name("Complete")

    if not complete_status_id:
        flash("Could not find a Complete status in task_status.", "error")
        return redirect(url_for("employee_tasks"))

    open_punches = fetch_all(
        """
        SELECT time_entry_ID, punch_in
        FROM time_punches
        WHERE task_id = ?
          AND punch_out IS NULL
        """,
        [task_id],
    )

    now_value = datetime.now()

    conn = get_connection()
    cur = conn.cursor()

    for punch in open_punches:
        minutes_worked = int((now_value - punch.punch_in).total_seconds() / 60)

        cur.execute(
            """
            UPDATE time_punches
            SET punch_out = ?,
                task_time = ?
            WHERE time_entry_ID = ?
            """,
            now_value,
            minutes_worked,
            punch.time_entry_ID,
        )

    cur.execute(
        """
        UPDATE tasks
        SET status_id = ?,
            actual_finish_date = ?
        WHERE task_id = ?
        """,
        complete_status_id,
        now_value,
        task_id,
    )

    conn.commit()
    conn.close()

    update_task_percentage_complete(task_id)

    flash("Task completed successfully. All open punches for this task were clocked out.", "success")
    return redirect(url_for("completion_photos", task_id=task_id))
########################################################################################################################
@app.route("/my-tasks")
@login_required
def employee_tasks():
    employee = get_logged_in_employee()

    if not employee:
        return redirect(url_for("login"))

    show_dept = request.args.get("show_dept") == "1"

    tasks = fetch_all(
        """
        SELECT
            t.task_id,
            t.task_name,
            t.task_description,
            t.planned_start_date,
            t.planned_finish_date,
            t.status_id,
            cc.coins_cost_code,
            cc.cost_code_description,
            li.line_item_description
        FROM
            (tasks AS t
            LEFT JOIN cost_codes AS cc
                ON t.cost_code_id = cc.cost_code_id)
            LEFT JOIN line_items AS li
                ON cc.line_item_id = li.line_item_id
        WHERE t.resource_id = ?
            AND t.status_id IN ('1', '2')
        ORDER BY t.planned_start_date, t.planned_finish_date, t.task_id
        """,
        [employee.resource_id],
    )

    dept_tasks = []
    dept_name = ""

    if show_dept and employee.resource_id:
        resource_row = fetch_one(
            "SELECT resource_types FROM resources WHERE resource_id = ?",
            [employee.resource_id],
        )

        if resource_row and resource_row.resource_types:
            resource_type_id = resource_row.resource_types

            rt_row = fetch_one(
                "SELECT resource_description FROM resource_types WHERE resource_type_id = ?",
                [resource_type_id],
            )
            dept_name = rt_row.resource_description if rt_row else ""

            dept_tasks = fetch_all(
                """
                SELECT
                    t.task_id,
                    t.task_name,
                    t.task_description,
                    t.planned_start_date,
                    t.planned_finish_date,
                    t.status_id,
                    cc.coins_cost_code,
                    li.line_item_description
                FROM
                    (tasks AS t
                    LEFT JOIN cost_codes AS cc
                        ON t.cost_code_id = cc.cost_code_id)
                    LEFT JOIN line_items AS li
                        ON cc.line_item_id = li.line_item_id
                WHERE t.resource_type_id = ?
                  AND t.status_id IN ('1', '2')
                ORDER BY t.planned_start_date, t.planned_finish_date, t.task_id
                """,
                [resource_type_id],
            )

    html = """
    <!doctype html>
    <html>
    <head>
        <title>My Tasks</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 40px; background: #f5f6f8; }
            .wrap { max-width: 1200px; margin: 0 auto; }
            .card { background: white; padding: 24px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,.08); margin-bottom: 20px; }
            table { width: 100%; border-collapse: collapse; }
            th, td { padding: 12px; border-bottom: 1px solid #ddd; text-align: left; vertical-align: top; font-size: 14px; }
            th { background: #fafafa; }
            .btn { padding: 10px 16px; background: #1f4e79; color: white; text-decoration: none; border: none; border-radius: 8px; font-weight: bold; cursor: pointer; display: inline-block; }
            .btn-secondary { background: #6c757d; }
            .btn-warning { background: #fd7e14; }
            .btn-green { background: #198754; }
            .top-bar { display: flex; justify-content: space-between; align-items: center; margin-bottom: 18px; }
            .button-row { display: flex; gap: 12px; flex-wrap: wrap; align-items: center; }
            .button-row form { margin: 0; }
            .desc-cell { max-width: 460px; white-space: normal; }
            .muted { color: #666; }
            .status-pill { display: inline-block; padding: 6px 10px; border-radius: 999px; background: #e9ecef; font-weight: bold; }
            .flash-success { background: #d1e7dd; color: #0f5132; padding: 10px; border-radius: 6px; margin-bottom: 12px; }
            .flash-error { background: #f8d7da; color: #842029; padding: 10px; border-radius: 6px; margin-bottom: 12px; }
        </style>
    </head>
    <body>
        <div class="wrap">
            <div class="top-bar">
                <div>
                    <h1>My Tasks</h1>
                    <div class="muted">Logged in as {{ employee.employee_first_name }} {{ employee.employee_last_name }}</div>
                </div>

                <div class="button-row">
                    {% if show_dept %}
                        <a class="btn btn-secondary" href="{{ url_for('employee_tasks') }}">Hide Department Tasks</a>
                    {% else %}
                        <a class="btn btn-green" href="{{ url_for('employee_tasks', show_dept=1) }}">All Department Tasks</a>
                    {% endif %}
                    <form method="post" action="{{ url_for('clock_in_task', task_id=tech_ops_task_id) }}">
                        <button class="btn" type="submit">Tech Ops</button>
                    </form>
                    <form method="post" action="{{ url_for('clock_in_task', task_id=overhead_task_id) }}">
                        <button class="btn btn-warning" type="submit">Overhead</button>
                    </form>
                    <form method="post" action="{{ url_for('employee_day_clock_out') }}">
                        <button class="btn btn-secondary" type="submit">Clock Out For Day</button>
                    </form>
                    <a class="btn" href="{{ url_for('logout') }}">Logout</a>
                </div>
            </div>

            {% with messages = get_flashed_messages(with_categories=true) %}
                {% if messages %}
                    {% for category, message in messages %}
                        <div class="flash-{{ category }}">{{ message }}</div>
                    {% endfor %}
                {% endif %}
            {% endwith %}

            <div class="card">
                <h2>Assigned Tasks</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Task</th>
                            <th>Description</th>
                            <th>Start</th>
                            <th>Finish</th>
                            <th>Status</th>
                            <th>Action</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for task in tasks %}
                        <tr>
                            <td><strong>{{ task.task_name or '' }}</strong></td>
                            <td class="desc-cell">
                                {{ task.line_item_description or task.task_description or '' }}
                            </td>
                            <td>{{ task.planned_start_date or '' }}</td>
                            <td>{{ task.planned_finish_date or '' }}</td>
                            <td>
                                <span class="status-pill">
                                    {% if task.status_id|int == 1 %}NOT STARTED{% elif task.status_id|int == 2 %}STARTED{% else %}UNKNOWN{% endif %}
                                </span>
                            </td>
                            <td>
                                <form method="post" action="{{ url_for('clock_in_task', task_id=task.task_id) }}">
                                    <button class="btn" type="submit">Clock In</button>
                                </form>
                            </td>
                        </tr>
                        {% else %}
                        <tr>
                            <td colspan="6">No tasks are currently assigned to you.</td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>

            {% if show_dept %}
            <div class="card">
                <h2>{{ dept_name }} — All Department Tasks</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Task</th>
                            <th>Description</th>
                            <th>Start</th>
                            <th>Finish</th>
                            <th>Status</th>
                            <th>Action</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for task in dept_tasks %}
                        <tr>
                            <td><strong>{{ task.task_name or '' }}</strong></td>
                            <td class="desc-cell">
                                {{ task.line_item_description or task.task_description or '' }}
                            </td>
                            <td>{{ task.planned_start_date or '' }}</td>
                            <td>{{ task.planned_finish_date or '' }}</td>
                            <td>
                                <span class="status-pill">
                                    {% if task.status_id|int == 1 %}NOT STARTED{% elif task.status_id|int == 2 %}STARTED{% else %}UNKNOWN{% endif %}
                                </span>
                            </td>
                            <td>
                                <form method="post" action="{{ url_for('clock_in_task', task_id=task.task_id) }}">
                                    <button class="btn" type="submit">Clock In</button>
                                </form>
                            </td>
                        </tr>
                        {% else %}
                        <tr>
                            <td colspan="6">No open tasks found for your department.</td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
            {% endif %}

        </div>
    </body>
    </html>
    """

    return render_template_string(
        html,
        employee=employee,
        tasks=tasks,
        dept_tasks=dept_tasks,
        dept_name=dept_name,
        show_dept=show_dept,
        tech_ops_task_id=TECH_OPS_TASK_ID,
        overhead_task_id=OVERHEAD_TASK_ID,
    )
########################################################################################################################
@app.route("/my-tasks/<int:task_id>")
@login_required
def employee_task_detail(task_id):
    employee = get_logged_in_employee()

    if not employee:
        return redirect(url_for("login"))

    # -------------------------------------------------------------------------
    # Get the employee's resource type / department.
    # -------------------------------------------------------------------------
    employee_resource = fetch_one(
        """
        SELECT resource_types
        FROM resources
        WHERE resource_id = ?
        """,
        [employee.resource_id],
    )

    employee_resource_type = (
        employee_resource.resource_types
        if employee_resource
        else None
    )

    # -------------------------------------------------------------------------
    # Load the task.
    #
    # Employee can open:
    #   1. A task assigned directly to their resource, OR
    #   2. Any task belonging to their department/resource type.
    # -------------------------------------------------------------------------
    task_row = fetch_one(
        """
        SELECT
            task_id,
            task_name,
            task_description,
            status_id,
            planned_hours,
            percentage_complete,
            resource_type_id,
            resource_id,
            cost_code_id,
            survey_type
        FROM tasks
        WHERE task_id = ?
          AND (
              resource_id = ?
              OR resource_type_id = ?
          )
        """,
        [
            task_id,
            employee.resource_id,
            employee_resource_type,
        ],
    )

    if not task_row:
        flash(
            "Task not found or not assigned to you.",
            "error",
        )
        return redirect(url_for("employee_tasks"))

    # -------------------------------------------------------------------------
    # Get cost code information separately.
    # -------------------------------------------------------------------------
    cost_code = None

    if task_row.cost_code_id:
        cost_code = fetch_one(
            """
            SELECT
                cost_code_id,
                coins_cost_code,
                cost_code_description,
                line_item_id
            FROM cost_codes
            WHERE cost_code_id = ?
            """,
            [task_row.cost_code_id],
        )

    # -------------------------------------------------------------------------
    # Get line item information separately.
    # -------------------------------------------------------------------------
    line_item = None

    if cost_code and cost_code.line_item_id:
        line_item = fetch_one(
            """
            SELECT
                line_item_id,
                line_item_description
            FROM line_items
            WHERE line_item_id = ?
            """,
            [cost_code.line_item_id],
        )

    # -------------------------------------------------------------------------
    # Combine everything into one object for the HTML.
    # -------------------------------------------------------------------------
    from types import SimpleNamespace

    task = SimpleNamespace(
        task_id=task_row.task_id,
        task_name=task_row.task_name,
        task_description=task_row.task_description,
        status_id=task_row.status_id,
        planned_hours=task_row.planned_hours,
        percentage_complete=task_row.percentage_complete,
        resource_type_id=task_row.resource_type_id,
        resource_id=task_row.resource_id,
        cost_code_id=task_row.cost_code_id,
        survey_type=task_row.survey_type,

        coins_cost_code=(
            cost_code.coins_cost_code
            if cost_code
            else ""
        ),

        cost_code_description=(
            cost_code.cost_code_description
            if cost_code
            else ""
        ),

        line_item_id=(
            line_item.line_item_id
            if line_item
            else None
        ),

        line_item_description=(
            line_item.line_item_description
            if line_item
            else ""
        ),
    )

    # -------------------------------------------------------------------------
    # Determine whether this employee is currently clocked into this task.
    # -------------------------------------------------------------------------
    open_punch = get_open_punch_for_employee(
        employee.employee_id
    )

    open_task_id = (
        open_punch.task_id
        if open_punch
        else None
    )

    # -------------------------------------------------------------------------
    # SCREEN
    # -------------------------------------------------------------------------
    html = """
    <!doctype html>
    <html>
    <head>
        <title>Task Detail</title>

        <style>
            body {
                font-family: Arial, sans-serif;
                margin: 40px;
                background: #f5f6f8;
            }

            .wrap {
                max-width: 900px;
                margin: 0 auto;
            }

            .card {
                background: white;
                padding: 32px;
                border-radius: 10px;
                box-shadow: 0 2px 10px rgba(0,0,0,.08);
                margin-bottom: 20px;
            }

            .btn {
                display: inline-block;
                padding: 14px 22px;
                background: #1f4e79;
                color: white;
                text-decoration: none;
                border: none;
                border-radius: 8px;
                font-weight: bold;
                cursor: pointer;
            }

            .big-btn {
                width: 100%;
                padding: 28px;
                font-size: 24px;
                margin-bottom: 14px;
                box-sizing: border-box;
                text-align: center;
            }

            .btn-secondary {
                background: #6c757d;
            }

            .btn-success {
                background: #198754;
            }

            .muted {
                color: #666;
            }

            .status-pill {
                display: inline-block;
                padding: 6px 10px;
                border-radius: 999px;
                background: #e9ecef;
                font-weight: bold;
            }

            .flash-success {
                background: #d1e7dd;
                color: #0f5132;
                padding: 10px;
                border-radius: 6px;
                margin-bottom: 12px;
            }

            .flash-error {
                background: #f8d7da;
                color: #842029;
                padding: 10px;
                border-radius: 6px;
                margin-bottom: 12px;
            }

            .info-row {
                margin-bottom: 14px;
            }

            .progress-wrap {
                position: relative;
                width: 100%;
                background: #e9ecef;
                border-radius: 999px;
                overflow: hidden;
                height: 38px;
                margin-top: 10px;
            }

            .progress-bar {
                height: 38px;
                background: #1f4e79;
                line-height: 38px;
            }

            .progress-text {
                position: absolute;
                left: 12px;
                top: 0;
                height: 38px;
                line-height: 38px;
                color: black;
                font-weight: bold;
                z-index: 5;
            }

            .progress-hours {
                position: absolute;
                right: 14px;
                top: 0;
                height: 38px;
                line-height: 38px;
                color: black;
                font-weight: bold;
                z-index: 5;
            }

            .progress-divider {
                position: absolute;
                top: 0;
                height: 38px;
                width: 2px;
                background: rgba(0,0,0,.35);
                z-index: 4;
            }

            form {
                margin: 0;
            }
        </style>
    </head>

    <body>
        <div class="wrap">

            {% with messages = get_flashed_messages(with_categories=true) %}

                {% if messages %}

                    {% for category, message in messages %}

                        <div class="flash-{{ category }}">
                            {{ message }}
                        </div>

                    {% endfor %}

                {% endif %}

            {% endwith %}


            <div class="card">

                <h1>
                    {{ task.task_name or 'Task' }}
                </h1>


                <div class="info-row">

                    <strong>Status:</strong>

                    <span class="status-pill">

                        {% if task.status_id|int == 1 %}
                            NOT STARTED

                        {% elif task.status_id|int == 2 %}
                            STARTED

                        {% elif task.status_id|int == 3 %}
                            COMPLETE

                        {% else %}
                            UNKNOWN

                        {% endif %}

                    </span>

                </div>


                <div class="info-row">

                    <strong>Description:</strong>
                    <br>

                    {{
                        task.line_item_description
                        or task.task_description
                        or ''
                    }}

                </div>


                <div class="info-row">

                    <strong>Cost Code:</strong>
                    <br>

                    {{
                        task.coins_cost_code
                        or task.cost_code_description
                        or ''
                    }}

                </div>


                {% if task.survey_type %}

                    <div class="info-row">

                        <strong>Survey Task</strong>
                        <br>

                        <span class="muted">
                            QC Form ID: {{ task.survey_type }}
                        </span>

                    </div>

                {% endif %}


                <div class="info-row">

                    <strong>Labor Budget Used:</strong>

                    {{
                        "%.1f"|format(
                            task.percentage_complete or 0
                        )
                    }}%

                    <div class="progress-wrap">

                        <div
                            class="progress-bar"
                            style="
                                width:
                                {{
                                    [
                                        task.percentage_complete or 0,
                                        100
                                    ]|min
                                }}%;
                            "
                        >
                        </div>


                        <div
                            class="progress-divider"
                            style="left:25%;"
                        ></div>

                        <div
                            class="progress-divider"
                            style="left:50%;"
                        ></div>

                        <div
                            class="progress-divider"
                            style="left:75%;"
                        ></div>


                        <div class="progress-text">

                            {{
                                "%.1f"|format(
                                    task.percentage_complete or 0
                                )
                            }}%

                        </div>


                        <div class="progress-hours">

                            {{
                                "%.2f"|format(
                                    task.planned_hours or 0
                                )
                            }}
                            hrs

                        </div>

                    </div>

                </div>

            </div>


            <div class="card">

                <a
                    class="btn big-btn"
                    href="{{
                        url_for(
                            'task_files',
                            task_id=task.task_id
                        )
                    }}"
                >
                    View Files
                </a>

            </div>


            {% if task.resource_type_id != installation_resource_type_id %}

                <div class="card">

                    <h2 style="color:red;">

                        <strong>
                            THIS SHOULD BE A BARCODE SCANNER
                            BUT YOU GET THE IDEA
                        </strong>

                    </h2>


                    <form
                        method="post"
                        action="{{
                            url_for(
                                'issue_material_to_task',
                                task_id=task.task_id
                            )
                        }}"
                    >

                        <div class="info-row">

                            <label>
                                <strong>
                                    Barcode
                                </strong>
                            </label>

                            <br>

                            <input
                                type="text"
                                name="barcode"
                                placeholder="Scan or enter barcode"
                                required
                                style="
                                    width:100%;
                                    padding:12px;
                                    font-size:18px;
                                    box-sizing:border-box;
                                "
                            >

                        </div>


                        <div class="info-row">

                            <label>
                                <strong>
                                    Quantity
                                </strong>
                            </label>

                            <br>

                            <input
                                type="number"
                                step="0.01"
                                min="0"
                                name="quantity"
                                required
                                style="
                                    width:100%;
                                    padding:12px;
                                    font-size:18px;
                                    box-sizing:border-box;
                                "
                            >

                        </div>


                        <button
                            class="btn big-btn"
                            type="submit"
                        >
                            Cost To Cost Code
                        </button>

                    </form>

                </div>

            {% endif %}


            <div class="card">

                <form
                    method="post"
                    action="{{
                        url_for(
                            'clock_out_task',
                            task_id=task.task_id
                        )
                    }}"
                >

                    <button
                        class="btn btn-secondary big-btn"
                        type="submit"
                    >
                        Clock Out
                    </button>

                </form>


                <form
                    method="post"
                    action="{{
                        url_for(
                            'complete_task',
                            task_id=task.task_id
                        )
                    }}"
                >

                    <button
                        class="btn btn-success big-btn"
                        type="submit"
                    >
                        Complete Task
                    </button>

                </form>

            </div>

        </div>
    </body>
    </html>
    """

    return render_template_string(
        html,
        employee=employee,
        task=task,
        open_task_id=open_task_id,
        installation_resource_type_id=INSTALLATION_RESOURCE_TYPE_ID,
    )
########################################################################################################################
@app.route("/my-tasks/<int:task_id>/issue-material", methods=["POST"])
@login_required
def issue_material_to_task(task_id):
    employee = get_logged_in_employee()

    if not employee:
        return redirect(url_for("login"))

    barcode = request.form.get("barcode", "").strip()
    quantity_text = request.form.get("quantity", "").strip()

    if not barcode or not quantity_text:
        flash("Barcode and quantity are required.", "error")
        return redirect(url_for("employee_task_detail", task_id=task_id))

    quantity = float(quantity_text)

    task = fetch_one(
        """
        SELECT
            task_id,
            resource_id,
            resource_type_id,
            cost_code_id
        FROM tasks
        WHERE task_id = ?
          AND (
              resource_id = ?
              OR resource_type_id = (
                  SELECT resource_types
                  FROM resources
                  WHERE resource_id = ?
              )
          )
        """,
        [task_id, employee.resource_id, employee.resource_id],
    )

    if not task:
        flash("Task not found or not assigned to you.", "error")
        return redirect(url_for("employee_tasks"))

    material = fetch_one(
        """
        SELECT material_id, material_name, barcode, unit_cost
        FROM materials
        WHERE barcode = ?
        """,
        [barcode],
    )

    if not material:
        flash(f"No inventory material found for barcode {barcode}.", "error")
        return redirect(url_for("employee_task_detail", task_id=task_id))

    transaction_date = datetime.now()

    conn = get_connection()
    cur = conn.cursor()

    cur.execute(
        """
        INSERT INTO material_transactions
            (material_id, cost_code_id, employee_id, transaction_date, quantity)
        VALUES (?, ?, ?, ?, ?)
        """,
        material.material_id,
        task.cost_code_id,
        employee.employee_id,
        transaction_date,
        quantity,
    )

    conn.commit()
    conn.close()

    flash(
        f"Costed {quantity} of {material.material_name} to this task.",
        "success"
    )

    return redirect(url_for("employee_task_detail", task_id=task_id))
########################################################################################################################
@app.route("/my-tasks/<int:task_id>/files")
@login_required
def task_files(task_id):
    employee = get_logged_in_employee()

    if not employee:
        return redirect(url_for("login"))

    task = fetch_one(
        """
        SELECT task_id
        FROM tasks
        WHERE task_id = ?
        """,
        [task_id],
    )

    if not task:
        flash("Task not found.", "error")
        return redirect(url_for("home"))

    possible_roots = [
        TASK_FILES_ROOT,
        r"C:\Users\Brooks\Desktop\Task Files",
        r"C:\Users\Brooks\OneDrive\Desktop\Task Files",
    ]

    folder_path = None

    for root in possible_roots:
        candidate = os.path.join(root, str(task_id))
        if os.path.exists(candidate):
            folder_path = candidate
            break

    if folder_path is None:
        folder_path = os.path.join(TASK_FILES_ROOT, str(task_id))
        os.makedirs(folder_path, exist_ok=True)

    print("TASK ID:", task_id)
    print("TASK FILES FOLDER:", folder_path)
    print("FOLDER EXISTS:", os.path.exists(folder_path))
    print("FILES:", os.listdir(folder_path) if os.path.exists(folder_path) else [])

    files = [
        f for f in os.listdir(folder_path)
        if os.path.isfile(os.path.join(folder_path, f))
    ]

    back_url = (
        url_for("office_task_detail", task_id=task_id)
        if getattr(employee, "office", False)
        else url_for("employee_task_detail", task_id=task_id)
    )

    html = """
    <!doctype html>
    <html>
    <head>
        <title>Task Files</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 40px; background: #f5f6f8; }
            .wrap { max-width: 900px; margin: 0 auto; }
            .card { background: white; padding: 32px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,.08); margin-bottom: 20px; }
            .btn { display: inline-block; padding: 14px 22px; background: #1f4e79; color: white; text-decoration: none; border: none; border-radius: 8px; font-weight: bold; cursor: pointer; margin-bottom: 8px; }
            .file-link { display: block; padding: 12px; border-bottom: 1px solid #ddd; color: #1f4e79; font-weight: bold; text-decoration: none; }
            .file-link:hover { background: #f8f9fa; }
            .muted { color: #666; font-size: 13px; }
        </style>
    </head>
    <body>
        <div class="wrap">
            <a class="btn" href="{{ back_url }}">← Back to Task</a>

            <div class="card">
                <h1>Task Files</h1>
                <p class="muted">Folder: {{ folder_path }}</p>

                {% if files %}
                    {% for file in files %}
                        <a class="file-link"
                           href="{{ url_for('download_task_file', task_id=task_id, filename=file) }}"
                           target="_blank">
                            {{ file }}
                        </a>
                    {% endfor %}
                {% else %}
                    <p>No files found for this task.</p>
                {% endif %}
            </div>
        </div>
    </body>
    </html>
    """

    return render_template_string(
        html,
        task_id=task_id,
        files=files,
        back_url=back_url,
        folder_path=folder_path,
    )
########################################################################################################################
@app.route("/my-tasks/<int:task_id>/files/<path:filename>")
@login_required
def download_task_file(task_id, filename):
    employee = get_logged_in_employee()

    if not employee:
        return redirect(url_for("login"))

    task = fetch_one(
        """
        SELECT task_id
        FROM tasks
        WHERE task_id = ?
        """,
        [task_id],
    )

    if not task:
        flash("Task not found.", "error")
        return redirect(url_for("home"))

    possible_roots = [
        TASK_FILES_ROOT,
        r"C:\Users\Brooks\Desktop\Task Files",
        r"C:\Users\Brooks\OneDrive\Desktop\Task Files",
    ]

    folder_path = None

    for root in possible_roots:
        candidate = os.path.join(root, str(task_id))
        if os.path.exists(candidate):
            folder_path = candidate
            break

    if folder_path is None:
        flash("Task folder not found.", "error")
        return redirect(url_for("task_files", task_id=task_id))

    return send_from_directory(folder_path, filename)
########################################################################################################################
@app.route("/my-tasks/<int:task_id>/indirect")
@login_required
def indirect_task_detail(task_id):
    employee = get_logged_in_employee()

    if not employee:
        return redirect(url_for("login"))

    if task_id not in INDIRECT_TASK_IDS:
        return redirect(url_for("employee_task_detail", task_id=task_id))

    task = fetch_one(
        """
        SELECT task_id, task_name
        FROM tasks
        WHERE task_id = ?
        """,
        [task_id],
    )

    if not task:
        flash("Task not found.", "error")
        return redirect(url_for("employee_tasks"))

    html = """
    <!doctype html>
    <html>
    <head>
        <title>{{ task.task_name }}</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 40px; background: #f5f6f8; }
            .wrap { max-width: 800px; margin: 0 auto; }
            .card { background: white; padding: 40px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,.08); text-align: center; }
            .btn { display: block; width: 100%; padding: 32px; background: #6c757d; color: white; border: none; border-radius: 8px; font-size: 26px; font-weight: bold; cursor: pointer; box-sizing: border-box; }
            h1 { font-size: 42px; }
        </style>
    </head>
    <body>
        <div class="wrap">
            <div class="card">
                <h1>{{ task.task_name }}</h1>
                <form method="post" action="{{ url_for('clock_out_task', task_id=task.task_id) }}">
                    <button class="btn" type="submit">CLOCK OUT OF {{ task.task_name.upper() }}</button>
                </form>
            </div>
        </div>
    </body>
    </html>
    """

    return render_template_string(html, task=task)
########################################################################################################################
@app.route("/job-wo/<int:task_id>")
@login_required
@office_required
def job_wo_detail(task_id):

    task = fetch_one(
        """
        SELECT
            li.job_number_id,
            li.work_order_id
        FROM
            (tasks AS t
            LEFT JOIN cost_codes AS cc
                ON t.cost_code_id = cc.cost_code_id)
            LEFT JOIN line_items AS li
                ON cc.line_item_id = li.line_item_id
        WHERE t.task_id = ?
        """,
        [task_id],
    )

    if not task:
        flash("Task not found.", "error")
        return redirect(url_for("home"))

    if task.job_number_id:

        html = """
        <!doctype html>
        <html>
        <head>
            <title>Construction Lab</title>
            <style>
                body {
                    font-family: Arial, sans-serif;
                    background: #f5f6f8;
                    text-align: center;
                    padding-top: 120px;
                }

                h1 {
                    font-size: 48px;
                    color: #1f4e79;
                }

                .btn {
                    display: inline-block;
                    padding: 14px 22px;
                    background: #1f4e79;
                    color: white;
                    text-decoration: none;
                    border-radius: 8px;
                    margin-top: 30px;
                }
            </style>
        </head>
        <body>

            <h1>THIS GOES TO CONSTRUCTION LAB</h1>

            <a class="btn" href="{{ url_for('home') }}">
                Back
            </a>

        </body>
        </html>
        """

        return render_template_string(html)

    elif task.work_order_id:

        html = """
        <!doctype html>
        <html>
        <head>
            <title>Service Lab</title>
            <style>
                body {
                    font-family: Arial, sans-serif;
                    background: #f5f6f8;
                    text-align: center;
                    padding-top: 120px;
                }

                h1 {
                    font-size: 48px;
                    color: #198754;
                }

                .btn {
                    display: inline-block;
                    padding: 14px 22px;
                    background: #198754;
                    color: white;
                    text-decoration: none;
                    border-radius: 8px;
                    margin-top: 30px;
                }
            </style>
        </head>
        <body>

            <h1>THIS GOES TO SERVICE LAB</h1>

            <a class="btn" href="{{ url_for('home') }}">
                Back
            </a>

        </body>
        </html>
        """

        return render_template_string(html)

    flash("No Job or Work Order linked to this task.", "error")
    return redirect(url_for("home"))
########################################################################################################################
@app.route("/my-tasks/<int:task_id>/completion-photos")
@login_required
def completion_photos(task_id):
    employee = get_logged_in_employee()

    if not employee:
        return redirect(url_for("login"))

    task = fetch_one(
        """
        SELECT task_id
        FROM tasks
        WHERE task_id = ?
        """,
        [task_id],
    )

    if not task:
        flash("Task not found.", "error")
        return redirect(url_for("employee_tasks"))

    html = """
    <!doctype html>
    <html>
    <head>
        <title>Completion Photos</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 40px; background: #f5f6f8; }
            .wrap { max-width: 800px; margin: 0 auto; }
            .card { background: white; padding: 40px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,.08); text-align: center; }
            .btn { display: block; width: 100%; padding: 28px; background: #1f4e79; color: white; text-decoration: none; border: none; border-radius: 8px; font-size: 24px; font-weight: bold; cursor: pointer; box-sizing: border-box; margin-top: 30px; }
            h1 { color: #1f4e79; }
            .demo-text { font-size: 24px; font-weight: bold; margin-top: 30px; line-height: 1.4; }
        </style>
    </head>
    <body>
        <div class="wrap">
            <div class="card">
                <h1>Completion Photos</h1>

                <div class="demo-text">
                    THIS IS FOR THE USER TO TAKE PHOTOS AND UPLOAD THEM TO THE TASK FOLDER.
                </div>

                <a class="btn" href="{{ url_for('quality_control_form', task_id=task_id) }}">
                    Upload Photos
                </a>
            </div>
        </div>
    </body>
    </html>
    """

    return render_template_string(html, task_id=task_id)
########################################################################################################################
@app.route("/my-tasks/<int:task_id>/quality-control", methods=["GET", "POST"])
@login_required
def quality_control_form(task_id):
    employee = get_logged_in_employee()

    if not employee:
        return redirect(url_for("login"))

    # =========================================================================
    # LOAD TASK
    # =========================================================================

    task = fetch_one(
        """
        SELECT
            t.task_id,
            t.task_name,
            t.resource_type_id,
            t.survey_type,
            li.job_number_id,
            li.work_order_id,
            li.line_item_description
        FROM ((tasks AS t
            INNER JOIN cost_codes AS cc
                ON t.cost_code_id = cc.cost_code_id)
            INNER JOIN line_items AS li
                ON cc.line_item_id = li.line_item_id)
        WHERE t.task_id = ?
        """,
        [task_id],
    )

    if not task:
        flash("Task not found.", "error")
        return redirect(url_for("employee_tasks"))

    # =========================================================================
    # JOB / WO LABEL
    # =========================================================================

    if task.job_number_id:
        job_wo_label = f"Job {task.job_number_id}"

    elif task.work_order_id:
        job_wo_label = f"WO {task.work_order_id}"

    else:
        job_wo_label = "Unknown"

    # =========================================================================
    # DETERMINE QC FORM
    #
    # Survey task:
    #     survey_type stores exact qc_form_id.
    #
    # Normal task:
    #     use resource_type_id.
    # =========================================================================

    if task.survey_type is not None:

        qc_form = fetch_one(
            """
            SELECT
                qc_form_id,
                form_name
            FROM qc_forms
            WHERE qc_form_id = ?
              AND active = True
            """,
            [task.survey_type],
        )

        if not qc_form:
            flash(
                "The survey form assigned to this task could not be found or is inactive.",
                "error",
            )
            return redirect(url_for("employee_tasks"))

    else:

        qc_form = fetch_one(
            """
            SELECT TOP 1
                qc_form_id,
                form_name
            FROM qc_forms
            WHERE resource_type_id = ?
              AND active = True
            ORDER BY qc_form_id
            """,
            [task.resource_type_id],
        )

        if not qc_form:
            flash(
                "No QC form found for this task's resource type.",
                "error",
            )
            return redirect(url_for("employee_tasks"))

    # =========================================================================
    # LOAD QUESTIONS
    # =========================================================================

    questions = fetch_all(
        """
        SELECT
            qc_question_id,
            question_text,
            question_type,
            required,
            sort_order,
            section_name,
            diagram_id
        FROM qc_form_questions
        WHERE qc_form_id = ?
          AND active = True
        ORDER BY
            sort_order,
            qc_question_id
        """,
        [qc_form.qc_form_id],
    )

    # =========================================================================
    # LOAD QUESTION OPTIONS
    # =========================================================================

    options_by_question = {}

    for question in questions:

        options = fetch_all(
            """
            SELECT
                qc_option_id,
                option_value,
                sort_order
            FROM qc_question_options
            WHERE qc_question_id = ?
              AND active = True
            ORDER BY
                sort_order,
                qc_option_id
            """,
            [question.qc_question_id],
        )

        options_by_question[question.qc_question_id] = options

    # =========================================================================
    # LOAD DIAGRAMS
    # =========================================================================

    diagrams = fetch_all(
        """
        SELECT
            diagram_id,
            diagram_title,
            image_path,
            sort_order
        FROM qc_form_diagram
        WHERE qc_form_id = ?
          AND active = True
        ORDER BY
            sort_order,
            diagram_id
        """,
        [qc_form.qc_form_id],
    )

    # =========================================================================
    # ORGANIZE QUESTIONS
    #
    # Questions with diagram_id = 0 / NULL are normal questions.
    # Questions with a diagram_id are displayed under that diagram.
    # =========================================================================

    normal_questions = []

    questions_by_diagram = {}

    for question in questions:

        diagram_id = question.diagram_id

        if diagram_id in (None, 0, "0", ""):
            normal_questions.append(question)

        else:
            try:
                diagram_key = int(diagram_id)
            except (TypeError, ValueError):
                diagram_key = diagram_id

            questions_by_diagram.setdefault(
                diagram_key,
                [],
            ).append(question)

    # =========================================================================
    # SAVE FORM
    # =========================================================================

    if request.method == "POST":

        submitted_date = datetime.now()

        conn = get_connection()
        cur = conn.cursor()

        for question in questions:

            field_name = (
                f"question_{question.qc_question_id}"
            )

            question_type = str(
                question.question_type or ""
            ).strip().lower()

            # -------------------------------------------------------------
            # Multi-select questions use checkboxes.
            # -------------------------------------------------------------

            if question_type == "multi_select":

                selected_values = request.form.getlist(
                    field_name
                )

                selected_values = [
                    value.strip()
                    for value in selected_values
                    if value.strip()
                ]

                answer_value = ", ".join(
                    selected_values
                )

            else:

                answer_value = request.form.get(
                    field_name,
                    "",
                ).strip()

            # -------------------------------------------------------------
            # Required validation
            # -------------------------------------------------------------

            if question.required and not answer_value:

                conn.close()

                flash(
                    f"Please answer: {question.question_text}",
                    "error",
                )

                return redirect(
                    url_for(
                        "quality_control_form",
                        task_id=task_id,
                    )
                )

            # -------------------------------------------------------------
            # Save answer
            # -------------------------------------------------------------

            cur.execute(
                """
                INSERT INTO qc_form_answers
                    (
                        task_id,
                        employee_id,
                        qc_question_id,
                        answer_value,
                        submitted_date
                    )
                VALUES (?, ?, ?, ?, ?)
                """,
                task_id,
                employee.employee_id,
                question.qc_question_id,
                answer_value,
                submitted_date,
            )

        # =========================================================================
        # ADDITIONAL NOTES
        # =========================================================================

        notes = request.form.get(
            "additional_notes",
            "",
        ).strip()

        if notes:

            cur.execute(
                """
                INSERT INTO qc_form_answers
                    (
                        task_id,
                        employee_id,
                        qc_question_id,
                        answer_value,
                        submitted_date
                    )
                VALUES (?, ?, ?, ?, ?)
                """,
                task_id,
                employee.employee_id,
                None,
                f"NOTES:{notes}",
                submitted_date,
            )

        conn.commit()
        conn.close()

        flash(
            f"{qc_form.form_name} submitted successfully.",
            "success",
        )

        return redirect(
            url_for("employee_tasks")
        )

    # =========================================================================
    # SCREEN
    # =========================================================================

    html = """
    <!doctype html>

    <html>

    <head>

        <title>
            {{ qc_form.form_name }}
        </title>

        <style>

            body {
                font-family: Arial, sans-serif;
                margin: 40px;
                background: #f5f6f8;
            }

            .wrap {
                max-width: 950px;
                margin: 0 auto;
            }

            .card {
                background: white;
                padding: 36px;
                border-radius: 10px;
                box-shadow: 0 2px 10px rgba(0,0,0,.08);
                margin-bottom: 20px;
            }

            h1 {
                color: #198754;
                font-size: 38px;
            }

            .section-title {
                font-size: 24px;
                font-weight: bold;
                color: #1f4e79;
                border-bottom: 3px solid #1f4e79;
                padding-bottom: 8px;
                margin-top: 35px;
                margin-bottom: 20px;
            }

            .question {
                margin-bottom: 24px;
                padding-bottom: 18px;
                border-bottom: 1px solid #ddd;
            }

            label {
                font-weight: bold;
                display: block;
                margin-bottom: 8px;
                font-size: 18px;
            }

            input,
            textarea,
            select {
                width: 100%;
                padding: 14px;
                font-size: 18px;
                box-sizing: border-box;
                border: 1px solid #ccc;
                border-radius: 8px;
            }

            textarea {
                min-height: 110px;
            }

            .required {
                color: red;
            }

            .btn {
                display: block;
                width: 100%;
                padding: 26px;
                background: #198754;
                color: white;
                text-decoration: none;
                border: none;
                border-radius: 8px;
                font-size: 24px;
                font-weight: bold;
                cursor: pointer;
                box-sizing: border-box;
                margin-top: 30px;
            }

            .survey-banner {
                background: #fff3cd;
                color: #664d03;
                padding: 10px 14px;
                border-radius: 7px;
                margin-bottom: 18px;
                font-weight: bold;
            }

            .diagram-card {
                background: #f8f9fa;
                padding: 22px;
                border: 1px solid #ddd;
                border-radius: 10px;
                margin: 30px 0;
            }

            .diagram-title {
                font-size: 24px;
                font-weight: bold;
                color: #1f4e79;
                margin-bottom: 16px;
            }

            .diagram-image {
                display: block;
                width: 100%;
                max-width: 800px;
                height: auto;
                margin: 0 auto 25px auto;
                border: 1px solid #ccc;
                border-radius: 6px;
                background: white;
            }

            .checkbox-grid {
                display: grid;
                grid-template-columns:
                    repeat(
                        auto-fit,
                        minmax(200px, 1fr)
                    );
                gap: 8px;
            }

            .checkbox-option {
                display: flex;
                align-items: center;
                gap: 8px;
                background: #f8f9fa;
                padding: 10px;
                border: 1px solid #ddd;
                border-radius: 6px;
                cursor: pointer;
                font-weight: normal;
                margin: 0;
            }

            .checkbox-option input {
                width: 18px;
                height: 18px;
                margin: 0;
                flex: 0 0 auto;
            }

            .flash-success {
                background: #d1e7dd;
                color: #0f5132;
                padding: 10px;
                border-radius: 6px;
                margin-bottom: 12px;
            }

            .flash-error {
                background: #f8d7da;
                color: #842029;
                padding: 10px;
                border-radius: 6px;
                margin-bottom: 12px;
            }

        </style>

    </head>


    <body>

        <div class="wrap">


            {% with messages = get_flashed_messages(with_categories=true) %}

                {% if messages %}

                    {% for category, message in messages %}

                        <div class="flash-{{ category }}">
                            {{ message }}
                        </div>

                    {% endfor %}

                {% endif %}

            {% endwith %}


            <div
                class="card"
                style="
                    background:#1f4e79;
                    color:white;
                    padding:20px 36px;
                "
            >

                <div
                    style="
                        font-size:13px;
                        opacity:.75;
                    "
                >

                    {% if task.survey_type %}
                        SURVEY FORM
                    {% else %}
                        QC FORM
                    {% endif %}

                </div>


                <div
                    style="
                        font-size:28px;
                        font-weight:bold;
                    "
                >

                    {{ job_wo_label }}

                </div>


                <div
                    style="
                        font-size:15px;
                        opacity:.85;
                        margin-top:4px;
                    "
                >

                    {{
                        task.line_item_description
                        or task.task_name
                        or ''
                    }}

                </div>

            </div>


            <div class="card">

                {% if task.survey_type %}

                    <div class="survey-banner">

                        Survey Task:
                        {{ task.task_name }}

                    </div>

                {% endif %}


                <h1>
                    {{ qc_form.form_name }}
                </h1>


                <form method="post">


                    {# ===================================================== #}
                    {# NORMAL QUESTIONS                                    #}
                    {# ===================================================== #}

                    {% set ns = namespace(last_section='') %}

                    {% for question in normal_questions %}

                        {% if question.section_name
                              and question.section_name != ns.last_section %}

                            <div class="section-title">
                                {{ question.section_name }}
                            </div>

                            {% set ns.last_section =
                                question.section_name %}

                        {% endif %}


                        <div class="question">

                            <label>

                                {{ question.question_text }}

                                {% if question.required %}

                                    <span class="required">
                                        *
                                    </span>

                                {% endif %}

                            </label>


                            {% if question.question_type == "yes_no" %}

                                <select
                                    name="question_{{ question.qc_question_id }}"
                                    {% if question.required %}
                                        required
                                    {% endif %}
                                >

                                    <option value="">
                                        -- Select --
                                    </option>

                                    <option value="Yes">
                                        Yes
                                    </option>

                                    <option value="No">
                                        No
                                    </option>

                                </select>


                            {% elif question.question_type == "pass_fail" %}

                                <select
                                    name="question_{{ question.qc_question_id }}"
                                    {% if question.required %}
                                        required
                                    {% endif %}
                                >

                                    <option value="">
                                        -- Select --
                                    </option>

                                    <option value="Pass">
                                        Pass
                                    </option>

                                    <option value="Fail">
                                        Fail
                                    </option>

                                </select>


                            {% elif question.question_type == "single_select" %}

                                <select
                                    name="question_{{ question.qc_question_id }}"
                                    {% if question.required %}
                                        required
                                    {% endif %}
                                >

                                    <option value="">
                                        -- Select --
                                    </option>

                                    {% for option in options_by_question.get(
                                        question.qc_question_id,
                                        []
                                    ) %}

                                        <option value="{{ option.option_value }}">
                                            {{ option.option_value }}
                                        </option>

                                    {% endfor %}

                                </select>


                            {% elif question.question_type == "multi_select" %}

                                <div class="checkbox-grid">

                                    {% for option in options_by_question.get(
                                        question.qc_question_id,
                                        []
                                    ) %}

                                        <label class="checkbox-option">

                                            <input
                                                type="checkbox"
                                                name="question_{{ question.qc_question_id }}"
                                                value="{{ option.option_value }}"
                                            >

                                            <span>
                                                {{ option.option_value }}
                                            </span>

                                        </label>

                                    {% endfor %}

                                </div>


                            {% elif question.question_type == "number" %}

                                <input
                                    type="number"
                                    step="0.125"
                                    name="question_{{ question.qc_question_id }}"
                                    {% if question.required %}
                                        required
                                    {% endif %}
                                >


                            {% elif question.question_type == "text" %}

                                <input
                                    type="text"
                                    name="question_{{ question.qc_question_id }}"
                                    {% if question.required %}
                                        required
                                    {% endif %}
                                >


                            {% else %}

                                <textarea
                                    name="question_{{ question.qc_question_id }}"
                                    {% if question.required %}
                                        required
                                    {% endif %}
                                ></textarea>

                            {% endif %}

                        </div>

                    {% endfor %}


                    {# ===================================================== #}
                    {# DIAGRAM QUESTIONS                                   #}
                    {# ===================================================== #}

                    {% for diagram in diagrams %}

                        <div class="diagram-card">

                            <div class="diagram-title">

                                {{ diagram.diagram_title }}

                            </div>


                            <img
                                class="diagram-image"
                                src="{{
                                    url_for(
                                        'qc_diagram_image',
                                        diagram_id=diagram.diagram_id
                                    )
                                }}"
                                alt="{{ diagram.diagram_title }}"
                            >


                            {% for question in questions_by_diagram.get(
                                diagram.diagram_id,
                                []
                            ) %}

                                <div class="question">

                                    <label>

                                        {{ question.question_text }}

                                        {% if question.required %}

                                            <span class="required">
                                                *
                                            </span>

                                        {% endif %}

                                    </label>


                                    {% if question.question_type == "yes_no" %}

                                        <select
                                            name="question_{{ question.qc_question_id }}"
                                            {% if question.required %}
                                                required
                                            {% endif %}
                                        >

                                            <option value="">
                                                -- Select --
                                            </option>

                                            <option value="Yes">
                                                Yes
                                            </option>

                                            <option value="No">
                                                No
                                            </option>

                                        </select>


                                    {% elif question.question_type == "single_select" %}

                                        <select
                                            name="question_{{ question.qc_question_id }}"
                                            {% if question.required %}
                                                required
                                            {% endif %}
                                        >

                                            <option value="">
                                                -- Select --
                                            </option>

                                            {% for option in options_by_question.get(
                                                question.qc_question_id,
                                                []
                                            ) %}

                                                <option value="{{ option.option_value }}">
                                                    {{ option.option_value }}
                                                </option>

                                            {% endfor %}

                                        </select>


                                    {% elif question.question_type == "multi_select" %}

                                        <div class="checkbox-grid">

                                            {% for option in options_by_question.get(
                                                question.qc_question_id,
                                                []
                                            ) %}

                                                <label class="checkbox-option">

                                                    <input
                                                        type="checkbox"
                                                        name="question_{{ question.qc_question_id }}"
                                                        value="{{ option.option_value }}"
                                                    >

                                                    <span>
                                                        {{ option.option_value }}
                                                    </span>

                                                </label>

                                            {% endfor %}

                                        </div>


                                    {% elif question.question_type == "number" %}

                                        <input
                                            type="number"
                                            step="0.125"
                                            name="question_{{ question.qc_question_id }}"
                                            {% if question.required %}
                                                required
                                            {% endif %}
                                        >


                                    {% elif question.question_type == "text" %}

                                        <input
                                            type="text"
                                            name="question_{{ question.qc_question_id }}"
                                            {% if question.required %}
                                                required
                                            {% endif %}
                                        >


                                    {% else %}

                                        <textarea
                                            name="question_{{ question.qc_question_id }}"
                                            {% if question.required %}
                                                required
                                            {% endif %}
                                        ></textarea>

                                    {% endif %}

                                </div>

                            {% endfor %}

                        </div>

                    {% endfor %}


                    {# ===================================================== #}
                    {# ADDITIONAL NOTES                                     #}
                    {# ===================================================== #}

                    <div class="question">

                        <label>
                            Additional Notes
                        </label>

                        <textarea
                            name="additional_notes"
                            placeholder="Enter any notes, observations, or issues here..."
                        ></textarea>

                    </div>


                    <button
                        class="btn"
                        type="submit"
                    >

                        Submit {{ qc_form.form_name }}

                    </button>

                </form>

            </div>

        </div>

    </body>

    </html>
    """

    return render_template_string(
        html,
        task_id=task_id,
        task=task,
        qc_form=qc_form,
        questions=questions,
        normal_questions=normal_questions,
        diagrams=diagrams,
        questions_by_diagram=questions_by_diagram,
        options_by_question=options_by_question,
        job_wo_label=job_wo_label,
    )
########################################################################################################################
@app.route("/qc-diagram/<int:diagram_id>")
@login_required
def qc_diagram_image(diagram_id):

    diagram = fetch_one(
        """
        SELECT
            diagram_id,
            image_path
        FROM qc_form_diagram
        WHERE diagram_id = ?
          AND active = True
        """,
        [diagram_id],
    )

    if not diagram:
        return "Diagram not found.", 404

    image_path = str(diagram.image_path or "").strip()

    if not image_path:
        return "Diagram image path is blank.", 404

    if not os.path.exists(image_path):
        return "Diagram image file not found.", 404

    return send_file(image_path)
########################################################################################################################
########################################################################################################################
########################################################################################################################





########################################################################################################################
########################################################################################################################
########################################################################################################################
"""MANUFACTURING SCHEDULING TEMPLATE MANAGEMENT"""
########################################################################################################################
########################################################################################################################
########################################################################################################################
@app.route("/templates", methods=["GET", "POST"])
@office_required
def manage_templates():
    conn = get_connection()
    cur = conn.cursor()

    if request.method == "POST":
        template_name = request.form.get("template_name", "").strip()
        template_description = request.form.get("template_description", "").strip()

        if not template_name:
            flash("Template name is required.", "error")
        else:
            cur.execute(
                """
                INSERT INTO tasking_templates (template_name, template_description)
                VALUES (?, ?)
                """,
                template_name,
                template_description,
            )
            conn.commit()
            flash("Template created successfully.", "success")
            conn.close()
            return redirect(url_for("manage_templates"))

    cur.execute(
        """
        SELECT template_id, template_name, template_description
        FROM tasking_templates
        ORDER BY template_name
        """
    )
    templates = cur.fetchall()
    conn.close()

    html = """
    <!doctype html>
    <html>
    <head>
        <title>Manage Tasking Templates</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 40px; background: #f5f6f8; }
            .wrap { max-width: 1000px; margin: 0 auto; }
            .card { background: white; padding: 24px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,.08); margin-bottom: 20px; }
            input[type='text'], textarea, select, input[type='number'] {
                width: 100%; padding: 10px; border: 1px solid #ccc; border-radius: 6px;
                margin-top: 6px; margin-bottom: 14px; box-sizing: border-box;
            }
            .btn { padding: 10px 16px; background: #1f4e79; color: white; text-decoration: none; border: none; border-radius: 8px; cursor: pointer; font-weight: bold; }
            table { width: 100%; border-collapse: collapse; }
            th, td { padding: 10px; border-bottom: 1px solid #ddd; text-align: left; vertical-align: top; }
            .flash-success { background: #d1e7dd; color: #0f5132; padding: 10px; border-radius: 6px; margin-bottom: 12px; }
            .flash-error { background: #f8d7da; color: #842029; padding: 10px; border-radius: 6px; margin-bottom: 12px; }
            .top-links a { text-decoration: none; margin-right: 10px; }
        </style>
    </head>
    <body>
        <div class="wrap">
            <div class="top-links" style="margin-bottom: 18px;">
                <a href="{{ url_for('home') }}">← Home</a>
            </div>

            {% with messages = get_flashed_messages(with_categories=true) %}
                {% if messages %}
                    {% for category, message in messages %}
                        <div class="flash-{{ category }}">{{ message }}</div>
                    {% endfor %}
                {% endif %}
            {% endwith %}

            <div class="card">
                <h2>Create Template</h2>
                <form method="post">
                    <label>Template Name</label>
                    <input type="text" name="template_name" required>

                    <label>Template Description</label>
                    <textarea name="template_description" rows="4"></textarea>

                    <button class="btn" type="submit">Create Template</button>
                </form>
            </div>

            <div class="card">
                <h2>Existing Templates</h2>
                <table>
                    <thead>
                        <tr>
                            <th>ID</th>
                            <th>Name</th>
                            <th>Description</th>
                            <th>Actions</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for t in templates %}
                        <tr>
                            <td>{{ t.template_id }}</td>
                            <td>{{ t.template_name }}</td>
                            <td>{{ t.template_description or '' }}</td>
                            <td>
                                <a class="btn" href="{{ url_for('edit_template', template_id=t.template_id) }}">Open / Edit</a>
                            </td>
                        </tr>
                        {% else %}
                        <tr>
                            <td colspan="4">No templates found.</td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </div>
    </body>
    </html>
    """
    return render_template_string(html, templates=templates)
########################################################################################################################
@app.route("/templates/<int:template_id>", methods=["GET", "POST"])
@office_required
def edit_template(template_id):
    conn = get_connection()
    cur = conn.cursor()

    action = request.form.get("action")

    if request.method == "POST":
        if action == "update_template":
            template_name = request.form.get("template_name", "").strip()
            template_description = request.form.get("template_description", "").strip()

            if not template_name:
                flash("Template name is required.", "error")
            else:
                cur.execute(
                    """
                    UPDATE tasking_templates
                    SET template_name = ?, template_description = ?
                    WHERE template_id = ?
                    """,
                    template_name,
                    template_description,
                    template_id,
                )
                conn.commit()
                flash("Template updated successfully.", "success")
                conn.close()
                return redirect(url_for("edit_template", template_id=template_id))

        elif action == "add_task":
            sequence_number = request.form.get("sequence_number", "").strip()
            resource_type_id = request.form.get("resource_type_id", "").strip()
            task_description = request.form.get("task_description", "").strip()

            if not sequence_number:
                flash("Sequence number is required.", "error")
            elif not resource_type_id:
                flash("Resource type is required.", "error")
            elif not task_description:
                flash("Task description is required.", "error")
            else:
                cur.execute(
                    """
                    INSERT INTO task_template_tasks (template_id, sequence_number, resource_type_id, task_description)
                    VALUES (?, ?, ?, ?)
                    """,
                    template_id,
                    int(sequence_number),
                    int(resource_type_id),
                    task_description,
                )
                conn.commit()
                flash("Template task added successfully.", "success")
                conn.close()
                return redirect(url_for("edit_template", template_id=template_id))

    cur.execute(
        """
        SELECT template_id, template_name, template_description
        FROM tasking_templates
        WHERE template_id = ?
        """,
        template_id,
    )
    template_row = cur.fetchone()

    if not template_row:
        conn.close()
        return "Template not found.", 404

    cur.execute(
        """
        SELECT 
            ttt.template_task_id,
            ttt.template_id,
            ttt.sequence_number,
            ttt.task_description,
            ttt.resource_type_id,
            rt.resource_description
        FROM task_template_tasks AS ttt
        LEFT JOIN resource_types AS rt
            ON ttt.resource_type_id = rt.resource_type_id
        WHERE ttt.template_id = ?
        ORDER BY ttt.sequence_number, ttt.template_task_id
        """,
        template_id,
    )
    tasks = cur.fetchall()

    resource_types = get_resource_types()
    conn.close()

    html = """
    <!doctype html>
    <html>
    <head>
        <title>Edit Template</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 40px; background: #f5f6f8; }
            .wrap { max-width: 1100px; margin: 0 auto; }
            .card { background: white; padding: 24px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,.08); margin-bottom: 20px; }
            input[type='text'], input[type='number'], textarea, select {
                width: 100%; padding: 10px; border: 1px solid #ccc; border-radius: 6px;
                margin-top: 6px; margin-bottom: 14px; box-sizing: border-box;
            }
            .btn { padding: 10px 16px; background: #1f4e79; color: white; text-decoration: none; border: none; border-radius: 8px; cursor: pointer; font-weight: bold; }
            table { width: 100%; border-collapse: collapse; }
            th, td { padding: 10px; border-bottom: 1px solid #ddd; text-align: left; }
            .flash-success { background: #d1e7dd; color: #0f5132; padding: 10px; border-radius: 6px; margin-bottom: 12px; }
            .flash-error { background: #f8d7da; color: #842029; padding: 10px; border-radius: 6px; margin-bottom: 12px; }
            .row { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
            .top-links a { text-decoration: none; margin-right: 10px; }
        </style>
    </head>
    <body>
        <div class="wrap">
            <div class="top-links" style="margin-bottom: 18px;">
                <a href="{{ url_for('home') }}">Home</a>
                <a href="{{ url_for('manage_templates') }}">← Back to Templates</a>
            </div>

            {% with messages = get_flashed_messages(with_categories=true) %}
                {% if messages %}
                    {% for category, message in messages %}
                        <div class="flash-{{ category }}">{{ message }}</div>
                    {% endfor %}
                {% endif %}
            {% endwith %}

            <div class="card">
                <h2>Edit Template Header</h2>
                <form method="post">
                    <input type="hidden" name="action" value="update_template">

                    <label>Template Name</label>
                    <input type="text" name="template_name" value="{{ template_row.template_name }}" required>

                    <label>Template Description</label>
                    <textarea name="template_description" rows="4">{{ template_row.template_description or '' }}</textarea>

                    <button class="btn" type="submit">Save Template</button>
                </form>
            </div>

            <div class="row">
                <div class="card">
                    <h2>Add Template Task</h2>
                    <form method="post">
                        <input type="hidden" name="action" value="add_task">

                        <label>Sequence Number</label>
                        <input type="number" name="sequence_number" required>

                        <label>Task Description</label>
                        <textarea name="task_description" rows="3" required></textarea>

                        <label>Resource Type</label>
                        <select name="resource_type_id" class="resource-type-select" required>
                            <option value="">-- Select Resource Type --</option>
                            {% for rt in resource_types %}
                                <option value="{{ rt.resource_type_id }}">{{ rt.resource_description }}</option>
                            {% endfor %}
                        </select>

                        <button class="btn" type="submit">Add Task</button>
                    </form>
                </div>

                <div class="card">
                    <h2>Current Task Sequence</h2>
                    <table>
                        <thead>
                            <tr>
                                <th>Template Task ID</th>
                                <th>Sequence</th>
                                <th>Description</th>
                                <th>Resource Type</th>
                            </tr>
                        </thead>
                        <tbody>
                            {% for task in tasks %}
                            <tr>
                                <td>{{ task.template_task_id }}</td>
                                <td>{{ task.sequence_number }}</td>
                                <td>{{ task.task_description or '' }}</td>
                                <td>{{ task.resource_description or '' }}</td>
                            </tr>
                            {% else %}
                            <tr>
                                <td colspan="4">No template tasks yet.</td>
                            </tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </body>
    </html>
    """

    return render_template_string(
        html,
        template_row=template_row,
        tasks=tasks,
        resource_types=resource_types,
    )
########################################################################################################################
########################################################################################################################
########################################################################################################################





########################################################################################################################
########################################################################################################################
########################################################################################################################

########################################################################################################################
@app.route("/manual-task", methods=["GET", "POST"])
@login_required
@office_required
def manual_task_assignment():
    """
    Emergency / one-day manual task assignment.

    The task is intentionally created WITHOUT:
      - sequence_number
      - planned_start_date
      - planned_finish_date

    Manufacturing schedulers ignore tasks with sequence_number IS NULL.
    """

    # -------------------------------------------------------------------------
    # Resources
    # -------------------------------------------------------------------------

    resources = fetch_all(
        """
        SELECT
            resource_id,
            resource_name,
            resource_types
        FROM resources
        ORDER BY resource_name
        """
    )

    # -------------------------------------------------------------------------
    # Jobs
    # -------------------------------------------------------------------------

    jobs = fetch_all(
        """
        SELECT
            j.job_id,
            c.customer_name
        FROM jobs AS j
        LEFT JOIN customers AS c
            ON j.customer_id = c.customer_id
        ORDER BY j.job_id DESC
        """
    )

    # -------------------------------------------------------------------------
    # Cost codes for all Jobs.
    #
    # Loaded once, then filtered in the browser based on selected Job.
    # -------------------------------------------------------------------------

    cost_codes = fetch_all(
        """
        SELECT
            cc.cost_code_id,
            cc.coins_cost_code,
            cc.cost_code_description,
            li.job_number_id,
            li.line_item_description
        FROM line_items AS li
        INNER JOIN cost_codes AS cc
            ON cc.line_item_id = li.line_item_id
        WHERE li.job_number_id IS NOT NULL
        ORDER BY
            li.job_number_id DESC,
            cc.coins_cost_code,
            cc.cost_code_description,
            cc.cost_code_id
        """
    )

    if request.method == "POST":

        job_id_raw = request.form.get(
            "job_id",
            "",
        ).strip()

        cost_code_id_raw = request.form.get(
            "cost_code_id",
            "",
        ).strip()

        resource_id_raw = request.form.get(
            "resource_id",
            "",
        ).strip()

        description = request.form.get(
            "task_description",
            "",
        ).strip()

        planned_hours_raw = request.form.get(
            "planned_hours",
            "",
        ).strip()

        # ---------------------------------------------------------------------
        # Validate basic entries
        # ---------------------------------------------------------------------

        if not job_id_raw:

            flash(
                "Select a Job.",
                "error",
            )

            return redirect(
                url_for(
                    "manual_task_assignment"
                )
            )

        if not cost_code_id_raw:

            flash(
                "Select a Cost Code.",
                "error",
            )

            return redirect(
                url_for(
                    "manual_task_assignment"
                )
            )

        if not resource_id_raw:

            flash(
                "Select a Resource.",
                "error",
            )

            return redirect(
                url_for(
                    "manual_task_assignment"
                )
            )

        if not description:

            flash(
                "Enter a task description.",
                "error",
            )

            return redirect(
                url_for(
                    "manual_task_assignment"
                )
            )

        if not planned_hours_raw:

            flash(
                "Enter the expected time in hours.",
                "error",
            )

            return redirect(
                url_for(
                    "manual_task_assignment"
                )
            )

        try:

            job_id = int(
                job_id_raw
            )

            cost_code_id = int(
                cost_code_id_raw
            )

            resource_id = int(
                resource_id_raw
            )

            planned_hours = float(
                planned_hours_raw
            )

        except (TypeError, ValueError):

            flash(
                "Job, Cost Code, Resource, and Hours must be valid.",
                "error",
            )

            return redirect(
                url_for(
                    "manual_task_assignment"
                )
            )

        if planned_hours <= 0:

            flash(
                "Expected time must be greater than zero.",
                "error",
            )

            return redirect(
                url_for(
                    "manual_task_assignment"
                )
            )

        # ---------------------------------------------------------------------
        # Server-side validation:
        # selected cost code MUST belong to selected Job.
        # ---------------------------------------------------------------------

        valid_cost_code = fetch_one(
            """
            SELECT
                cc.cost_code_id
            FROM line_items AS li
            INNER JOIN cost_codes AS cc
                ON cc.line_item_id = li.line_item_id
            WHERE li.job_number_id = ?
              AND cc.cost_code_id = ?
            """,
            [
                job_id,
                cost_code_id,
            ],
        )

        if not valid_cost_code:

            flash(
                "That Cost Code does not belong to the selected Job.",
                "error",
            )

            return redirect(
                url_for(
                    "manual_task_assignment"
                )
            )

        # ---------------------------------------------------------------------
        # Get the selected resource and automatically inherit its resource type.
        # ---------------------------------------------------------------------

        resource = fetch_one(
            """
            SELECT
                resource_id,
                resource_name,
                resource_types
            FROM resources
            WHERE resource_id = ?
            """,
            [resource_id],
        )

        if not resource:

            flash(
                "Selected Resource was not found.",
                "error",
            )

            return redirect(
                url_for(
                    "manual_task_assignment"
                )
            )

        resource_type_id = None

        if resource.resource_types not in (
            None,
            "",
        ):

            try:
                resource_type_id = int(
                    resource.resource_types
                )

            except (TypeError, ValueError):

                flash(
                    "The selected Resource does not have a valid Resource Type.",
                    "error",
                )

                return redirect(
                    url_for(
                        "manual_task_assignment"
                    )
                )

        status_id = get_status_id_by_name(
            DEFAULT_NOT_STARTED_STATUS_NAME
        )

        if not status_id:

            flash(
                "Could not find the Not Started task status.",
                "error",
            )

            return redirect(
                url_for(
                    "manual_task_assignment"
                )
            )

        # ---------------------------------------------------------------------
        # Create the emergency task.
        #
        # sequence_number is deliberately NULL.
        # planned_start_date / planned_finish_date are deliberately NULL.
        #
        # This means:
        #   - it appears immediately in My Tasks for the assigned resource
        #   - it does not appear on the Gantt chart
        #   - manufacturing scheduling algorithms ignore it
        # ---------------------------------------------------------------------

        conn = get_connection()
        cur = conn.cursor()

        cur.execute(
            """
            INSERT INTO tasks
                (
                    cost_code_id,
                    task_name,
                    task_description,
                    sequence_number,
                    planned_hours,
                    resource_type_id,
                    resource_id,
                    status_id,
                    percentage_complete,
                    planned_start_date,
                    planned_finish_date
                )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            cost_code_id,
            f"Manual Task — Job {job_id}",
            description,
            None,
            planned_hours,
            resource_type_id,
            resource_id,
            status_id,
            0,
            None,
            None,
        )

        conn.commit()
        conn.close()

        flash(
            (
                f"Manual task assigned to "
                f"{resource.resource_name}."
            ),
            "success",
        )

        return redirect(
            url_for(
                "manual_task_assignment"
            )
        )

    # -------------------------------------------------------------------------
    # Screen
    # -------------------------------------------------------------------------

    html = """
    <!doctype html>

    <html>

    <head>

        <title>
            Assign Manual Task
        </title>

        <style>

            * {
                box-sizing: border-box;
            }

            body {
                font-family: Arial, sans-serif;
                margin: 0;
                padding: 30px;
                background: #f5f6f8;
            }

            .wrap {
                max-width: 900px;
                margin: auto;
            }

            .card {
                background: white;
                padding: 28px;
                border-radius: 10px;
                box-shadow:
                    0 2px 10px
                    rgba(0,0,0,.08);
                margin-bottom: 20px;
            }

            h1 {
                color: #1f4e79;
                margin-top: 0;
            }

            .note {
                padding: 14px;
                background: #fff3cd;
                border-radius: 7px;
                margin-bottom: 20px;
                line-height: 1.5;
            }

            .grid {
                display: grid;
                grid-template-columns:
                    1fr 1fr;
                gap: 18px;
            }

            label {
                display: block;
                font-weight: bold;
                margin-bottom: 6px;
            }

            input,
            select,
            textarea {
                width: 100%;
                padding: 11px;
                border: 1px solid #ccc;
                border-radius: 6px;
                margin-bottom: 16px;
                font-size: 14px;
            }

            textarea {
                min-height: 120px;
                resize: vertical;
            }

            .btn {
                display: inline-block;
                padding: 11px 18px;
                border: none;
                border-radius: 7px;
                background: #1f4e79;
                color: white;
                text-decoration: none;
                font-weight: bold;
                cursor: pointer;
            }

            .btn-back {
                background: #6c757d;
                margin-bottom: 18px;
            }

            .flash-success {
                background: #d1e7dd;
                color: #0f5132;
                padding: 10px;
                border-radius: 6px;
                margin-bottom: 15px;
            }

            .flash-error {
                background: #f8d7da;
                color: #842029;
                padding: 10px;
                border-radius: 6px;
                margin-bottom: 15px;
            }

            @media(max-width:700px) {

                .grid {
                    grid-template-columns: 1fr;
                }

            }

        </style>


        <script>

            function filterCostCodes() {

                const jobId =
                    document.getElementById(
                        "job_id"
                    ).value;

                const costCodeSelect =
                    document.getElementById(
                        "cost_code_id"
                    );

                costCodeSelect.value = "";

                const options =
                    costCodeSelect.querySelectorAll(
                        "option[data-job-id]"
                    );

                options.forEach(
                    function(option) {

                        option.hidden =
                            option.dataset.jobId
                            !== jobId;

                    }
                );
            }


            window.addEventListener(
                "DOMContentLoaded",
                function() {

                    filterCostCodes();

                }
            );

        </script>

    </head>


    <body>

        <div class="wrap">

            <a
                class="btn btn-back"
                href="{{ url_for('home') }}"
            >
                ← Home
            </a>


            {% with messages =
                get_flashed_messages(
                    with_categories=true
                )
            %}

                {% if messages %}

                    {% for category, message
                        in messages
                    %}

                        <div
                            class="flash-{{ category }}"
                        >
                            {{ message }}
                        </div>

                    {% endfor %}

                {% endif %}

            {% endwith %}


            <div class="card">

                <h1>
                    Assign Manual Task
                </h1>

                <div class="note">

                    Use this for emergency or same-day work.
                    The task is assigned directly to the selected
                    resource and is intentionally ignored by the
                    manufacturing scheduling algorithm.

                </div>


                <form method="post">

                    <div class="grid">

                        <div>

                            <label>
                                Resource
                            </label>

                            <select
                                name="resource_id"
                                required
                            >

                                <option value="">
                                    -- Select Resource --
                                </option>

                                {% for resource in resources %}

                                    <option
                                        value="{{ resource.resource_id }}"
                                    >
                                        {{ resource.resource_name }}
                                    </option>

                                {% endfor %}

                            </select>

                        </div>


                        <div>

                            <label>
                                Expected Time (Hours)
                            </label>

                            <input
                                type="number"
                                name="planned_hours"
                                min="0.1"
                                step="0.1"
                                placeholder="Example: 2.5"
                                required
                            >

                        </div>

                    </div>


                    <div class="grid">

                        <div>

                            <label>
                                Job
                            </label>

                            <select
                                id="job_id"
                                name="job_id"
                                onchange="filterCostCodes()"
                                required
                            >

                                <option value="">
                                    -- Select Job --
                                </option>

                                {% for job in jobs %}

                                    <option
                                        value="{{ job.job_id }}"
                                    >
                                        Job {{ job.job_id }}

                                        {% if job.customer_name %}
                                            — {{ job.customer_name }}
                                        {% endif %}
                                    </option>

                                {% endfor %}

                            </select>

                        </div>


                        <div>

                            <label>
                                Cost Code
                            </label>

                            <select
                                id="cost_code_id"
                                name="cost_code_id"
                                required
                            >

                                <option value="">
                                    -- Select Job First --
                                </option>

                                {% for cc in cost_codes %}

                                    <option
                                        value="{{ cc.cost_code_id }}"
                                        data-job-id="{{ cc.job_number_id }}"
                                    >

                                        {{
                                            cc.coins_cost_code
                                            or cc.cost_code_id
                                        }}

                                        {% if cc.cost_code_description %}
                                            — {{ cc.cost_code_description }}
                                        {% endif %}

                                        {% if cc.line_item_description %}
                                            — {{ cc.line_item_description }}
                                        {% endif %}

                                    </option>

                                {% endfor %}

                            </select>

                        </div>

                    </div>


                    <label>
                        Description
                    </label>

                    <textarea
                        name="task_description"
                        placeholder="What needs to be done?"
                        required
                    ></textarea>


                    <button
                        class="btn"
                        type="submit"
                    >
                        Assign Task
                    </button>

                </form>

            </div>

        </div>

    </body>

    </html>
    """

    return render_template_string(
        html,
        resources=resources,
        jobs=jobs,
        cost_codes=cost_codes,
    )
########################################################################################################################

"""SCHEDULING PIPELINE"""
########################################################################################################################
########################################################################################################################
########################################################################################################################
@app.route("/schedule", methods=["GET"])
@office_required
def schedule_search():
    search_text = request.args.get("q", "").strip()
    like_text = f"%{search_text}%"

    if search_text:
        jobs = fetch_all(
            """
            SELECT TOP 100
                j.job_id,
                j.customer_id,
                c.customer_name,
                j.due_date
            FROM jobs AS j
            LEFT JOIN customers AS c
                ON j.customer_id = c.customer_id
            WHERE CStr(j.job_id) LIKE ?
               OR c.customer_name LIKE ?
            ORDER BY j.job_id DESC
            """,
            [like_text, like_text],
        )

        work_orders = fetch_all(
            """
            SELECT TOP 100
                w.wo_id,
                w.customer_id,
                c.customer_name,
                w.due_date
            FROM work_orders AS w
            LEFT JOIN customers AS c
                ON w.customer_id = c.customer_id
            WHERE CStr(w.wo_id) LIKE ?
               OR c.customer_name LIKE ?
            ORDER BY w.wo_id DESC
            """,
            [like_text, like_text],
        )
    else:
        jobs = fetch_all(
            """
            SELECT TOP 100
                j.job_id,
                j.customer_id,
                c.customer_name,
                j.due_date
            FROM jobs AS j
            LEFT JOIN customers AS c
                ON j.customer_id = c.customer_id
            ORDER BY j.job_id DESC
            """
        )

        work_orders = fetch_all(
            """
            SELECT TOP 100
                w.wo_id,
                w.customer_id,
                c.customer_name,
                w.due_date
            FROM work_orders AS w
            LEFT JOIN customers AS c
                ON w.customer_id = c.customer_id
            ORDER BY w.wo_id DESC
            """
        )

    html = """
    <!doctype html>
    <html>
    <head>
        <title>Schedule Job / WO</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 40px; background: #f5f6f8; }
            .wrap { max-width: 1200px; margin: 0 auto; }
            .card { background: white; padding: 24px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,.08); margin-bottom: 20px; }
            input[type='text'] { width: 100%; padding: 10px; border: 1px solid #ccc; border-radius: 6px; margin: 8px 0 14px 0; box-sizing: border-box; }
            .btn { padding: 10px 16px; background: #1f4e79; color: white; text-decoration: none; border: none; border-radius: 8px; cursor: pointer; font-weight: bold; }
            table { width: 100%; border-collapse: collapse; }
            th, td { padding: 10px; border-bottom: 1px solid #ddd; text-align: left; }
            .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
            .top-links a { text-decoration: none; margin-right: 10px; }
        </style>
    </head>
    <body>
        <div class="wrap">
            <div class="top-links" style="margin-bottom: 18px;">
                <a href="{{ url_for('home') }}">← Home</a>
            </div>

            <div class="card">
                <h2>Schedule Job / WO</h2>
                <form method="get">
                    <label>Search by Job ID, WO ID, or Customer Name</label>
                    <input type="text" name="q" value="{{ search_text }}">
                    <button class="btn" type="submit">Search</button>
                </form>

                <hr style="margin:22px 0; border:none; border-top:1px solid #ddd;">

                <form
                    method="post"
                    action="{{ url_for('resequence_schedule') }}"
                    onsubmit="return confirm('Globally re-optimize every movable manufacturing task? Started and completed tasks will not move.');"
                >
                    <button
                        class="btn"
                        type="submit"
                        style="background:#198754;"
                    >
                        Re-Optimize Entire Manufacturing Schedule
                    </button>
                </form>

                <p style="color:#666; font-size:13px;">
                    All Jobs and Work Orders compete for shared manufacturing resources using due date, customer importance, lateness, and final-task priority.
                </p>
            </div>

            <div class="grid">
                <div class="card">
                    <h3>Jobs</h3>
                    <table>
                        <thead>
                            <tr>
                                <th>Job ID</th>
                                <th>Customer</th>
                                <th>Due Date</th>
                                <th></th>
                            </tr>
                        </thead>
                        <tbody>
                            {% for row in jobs %}
                            <tr>
                                <td>{{ row.job_id }}</td>
                                <td>{{ row.customer_name or '' }}</td>
                                <td>{{ row.due_date or '' }}</td>
                                <td>
                                    <a class="btn" href="{{ url_for('job_line_items', entity_type='job', entity_id=row.job_id) }}">Open</a>
                                </td>
                            </tr>
                            {% else %}
                            <tr><td colspan="4">No jobs found.</td></tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>

                <div class="card">
                    <h3>Work Orders</h3>
                    <table>
                        <thead>
                            <tr>
                                <th>WO ID</th>
                                <th>Customer</th>
                                <th>Due Date</th>
                                <th></th>
                            </tr>
                        </thead>
                        <tbody>
                            {% for row in work_orders %}
                            <tr>
                                <td>{{ row.wo_id }}</td>
                                <td>{{ row.customer_name or '' }}</td>
                                <td>{{ row.due_date or '' }}</td>
                                <td>
                                    <a class="btn" href="{{ url_for('job_line_items', entity_type='wo', entity_id=row.wo_id) }}">Open</a>
                                </td>
                            </tr>
                            {% else %}
                            <tr><td colspan="4">No work orders found.</td></tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </body>
    </html>
    """
    return render_template_string(html, jobs=jobs, work_orders=work_orders, search_text=search_text)
########################################################################################################################
@app.route("/schedule/<entity_type>/<int:entity_id>/line-items", methods=["GET", "POST"])
@office_required
def job_line_items(entity_type, entity_id):
    if request.method == "POST":
        result = run_global_scheduler()
        if result["scheduled"] > 0:
            _mark_manufacturing_scheduled(
                entity_type,
                entity_id,
            )

        flash(
            (
                f"{result['message']} "
                f"Scheduled: {result['scheduled']} | "
                f"Skipped: {result.get('skipped', 0)} | "
                f"Late Tasks: {result['late_tasks']} | "
                f"Late Final Tasks: {result['late_final_tasks']} | "
                f"Total Loss: {result['loss']:.2f}"
            ),
            "success" if result["scheduled"] > 0 else "error",
        )
        return redirect(url_for("job_line_items", entity_type=entity_type, entity_id=entity_id))

    if entity_type == "job":
        try:
            header = fetch_one(
                """
                SELECT
                    j.job_id AS entity_id,
                    c.customer_name,
                    j.due_date,
                    j.manufacturing_due_date
                FROM jobs AS j
                LEFT JOIN customers AS c
                    ON j.customer_id = c.customer_id
                WHERE j.job_id = ?
                """,
                [entity_id],
            )
        except Exception:
            header = fetch_one(
                """
                SELECT
                    j.job_id AS entity_id,
                    c.customer_name,
                    j.due_date
                FROM jobs AS j
                LEFT JOIN customers AS c
                    ON j.customer_id = c.customer_id
                WHERE j.job_id = ?
                """,
                [entity_id],
            )

        line_items = fetch_all(
            """
            SELECT line_item_id, job_number_id, work_order_id, quantity, line_item_description
            FROM line_items
            WHERE job_number_id = ?
            ORDER BY line_item_id
            """,
            [entity_id],
        )
        header_label = "Job"
    else:
        try:
            header = fetch_one(
                """
                SELECT
                    w.wo_id AS entity_id,
                    c.customer_name,
                    w.due_date,
                    w.manufacturing_due_date
                FROM work_orders AS w
                LEFT JOIN customers AS c
                    ON w.customer_id = c.customer_id
                WHERE w.wo_id = ?
                """,
                [entity_id],
            )
        except Exception:
            header = fetch_one(
                """
                SELECT
                    w.wo_id AS entity_id,
                    c.customer_name,
                    w.due_date
                FROM work_orders AS w
                LEFT JOIN customers AS c
                    ON w.customer_id = c.customer_id
                WHERE w.wo_id = ?
                """,
                [entity_id],
            )

        line_items = fetch_all(
            """
            SELECT line_item_id, job_number_id, work_order_id, quantity, line_item_description
            FROM line_items
            WHERE work_order_id = ?
            ORDER BY line_item_id
            """,
            [entity_id],
        )
        header_label = "Work Order"

    if not header:
        return "Record not found.", 404

    html = """
    <!doctype html>
    <html>
    <head>
        <title>Line Items</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 40px; background: #f5f6f8; }
            .wrap { max-width: 1100px; margin: 0 auto; }
            .card { background: white; padding: 24px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,.08); margin-bottom: 20px; }
            .btn { padding: 10px 16px; background: #1f4e79; color: white; text-decoration: none; border: none; border-radius: 8px; cursor: pointer; font-weight: bold; }
            table { width: 100%; border-collapse: collapse; }
            th, td { padding: 10px; border-bottom: 1px solid #ddd; text-align: left; vertical-align: top; }
            .top-links a { text-decoration: none; margin-right: 10px; }
            .desc-cell { max-width: 500px; white-space: normal; }
            .flash-success { background: #d1e7dd; color: #0f5132; padding: 10px; border-radius: 6px; margin-bottom: 12px; }
            .flash-error { background: #f8d7da; color: #842029; padding: 10px; border-radius: 6px; margin-bottom: 12px; }
        </style>
    </head>
    <body>
        <div class="wrap">
            <div class="top-links" style="margin-bottom: 18px;">
                <a href="{{ url_for('home') }}">Home</a>
                <a href="{{ url_for('schedule_search') }}">← Back to Search</a>
            </div>

            {% with messages = get_flashed_messages(with_categories=true) %}
                {% if messages %}
                    {% for category, message in messages %}
                        <div class="flash-{{ category }}">{{ message }}</div>
                    {% endfor %}
                {% endif %}
            {% endwith %}

            <div class="card">
                <h2>{{ header_label }} {{ header.entity_id }}</h2>
                <p><strong>Customer:</strong> {{ header.customer_name or '' }}</p>
                <p><strong>Project Due Date:</strong> {{ header.due_date or '' }}</p>

                <p>
                    <strong>Manufacturing Due Date:</strong>
                    {% if header.manufacturing_due_date is defined and header.manufacturing_due_date %}
                        {{ header.manufacturing_due_date }}
                    {% else %}
                        Not Set — project due date will be used
                    {% endif %}
                </p>

                <form
                    method="post"
                    action="{{
                        url_for(
                            'set_manufacturing_due_date',
                            entity_type=entity_type,
                            entity_id=entity_id
                        )
                    }}"
                    style="margin-bottom:16px;"
                >
                    <input
                        type="date"
                        name="manufacturing_due_date"
                        value="{% if header.manufacturing_due_date is defined and header.manufacturing_due_date %}{{ header.manufacturing_due_date.strftime('%Y-%m-%d') }}{% endif %}"
                        required
                        style="padding:9px; border:1px solid #ccc; border-radius:6px;"
                    >

                    <button
                        class="btn"
                        type="submit"
                    >
                        Set Manufacturing Due Date
                    </button>
                </form>

                <form
                    method="post"
                    onsubmit="return confirm('This will globally re-optimize all movable manufacturing tasks across all Jobs and Work Orders. Started and completed tasks will not move. Continue?');"
                >
                    <button
                        class="btn"
                        type="submit"
                    >
                        Schedule / Re-Optimize Manufacturing
                    </button>
                </form>

                <p style="color:#666; font-size:13px; margin-top:10px;">
                    Scheduling is global: all movable manufacturing work competes for shared department capacity.
                </p>
            </div>

            <div class="card">
                <h2>Line Items</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Line Item ID</th>
                            <th>Description</th>
                            <th>Quantity</th>
                            <th></th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for row in line_items %}
                        <tr>
                            <td>{{ row.line_item_id }}</td>
                            <td class="desc-cell">{{ row.line_item_description or '' }}</td>
                            <td>{{ row.quantity or '' }}</td>
                            <td>
                                <a class="btn" href="{{ url_for('line_item_cost_codes', entity_type=entity_type, entity_id=entity_id, line_item_id=row.line_item_id) }}">Open Cost Codes</a>
                            </td>
                        </tr>
                        {% else %}
                        <tr><td colspan="4">No line items found.</td></tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </div>
    </body>
    </html>
    """
    return render_template_string(html, header=header, header_label=header_label, line_items=line_items, entity_type=entity_type, entity_id=entity_id)
########################################################################################################################
@app.route("/<entity_type>/<int:entity_id>/set-manufacturing-due-date", methods=["POST"])
@office_required
def set_manufacturing_due_date(entity_type, entity_id):

    value = request.form.get(
        "manufacturing_due_date",
        "",
    ).strip()

    if not value:

        flash(
            "Manufacturing due date is required.",
            "error",
        )

        return redirect(
            url_for(
                "job_line_items",
                entity_type=entity_type,
                entity_id=entity_id,
            )
        )

    try:
        due_dt = datetime.strptime(
            value,
            "%Y-%m-%d",
        )

    except ValueError:

        flash(
            "Invalid manufacturing due date.",
            "error",
        )

        return redirect(
            url_for(
                "job_line_items",
                entity_type=entity_type,
                entity_id=entity_id,
            )
        )

    try:

        if entity_type == "job":

            execute_sql(
                """
                UPDATE jobs
                SET manufacturing_due_date = ?
                WHERE job_id = ?
                """,
                [
                    due_dt,
                    entity_id,
                ],
            )

        elif entity_type == "wo":

            execute_sql(
                """
                UPDATE work_orders
                SET manufacturing_due_date = ?
                WHERE wo_id = ?
                """,
                [
                    due_dt,
                    entity_id,
                ],
            )

        else:

            flash(
                "Invalid record type.",
                "error",
            )

            return redirect(
                url_for("home")
            )

    except Exception as exc:

        flash(
            (
                "Could not save manufacturing due date. "
                "Make sure this table has a manufacturing_due_date field. "
                f"Database message: {exc}"
            ),
            "error",
        )

        return redirect(
            url_for(
                "job_line_items",
                entity_type=entity_type,
                entity_id=entity_id,
            )
        )

    flash(
        "Manufacturing due date updated.",
        "success",
    )

    return redirect(
        url_for(
            "job_line_items",
            entity_type=entity_type,
            entity_id=entity_id,
        )
    )
########################################################################################################################
@app.route("/schedule/resequence", methods=["POST"])
@office_required
def resequence_schedule():

    result = run_global_scheduler()

    flash(
        (
            f"{result['message']} "
            f"Scheduled: {result['scheduled']} | "
            f"Skipped: {result.get('skipped', 0)} | "
            f"Late Tasks: {result['late_tasks']} | "
            f"Late Final Tasks: {result['late_final_tasks']} | "
            f"Total Loss: {result['loss']:.2f}"
        ),
        "success"
        if result["scheduled"] > 0
        else "error",
    )

    return redirect(
        url_for("schedule_search")
    )
########################################################################################################################
########################################################################################################################
@app.route("/schedule/<entity_type>/<int:entity_id>/line-items/<int:line_item_id>/cost-codes")
@office_required
def line_item_cost_codes(entity_type, entity_id, line_item_id):
    cost_codes = fetch_all(
        """
        SELECT cc.cost_code_id, cc.cost_code_description, cc.coins_cost_code
        FROM cost_codes AS cc
        WHERE cc.line_item_id = ?
        ORDER BY cc.cost_code_id
        """,
        [line_item_id],
    )

    cost_code_rows = []
    for row in cost_codes:
        task_count_row = fetch_one(
            "SELECT COUNT(*) AS task_count FROM tasks WHERE cost_code_id = ?",
            [row.cost_code_id],
        )
        task_count = task_count_row.task_count if task_count_row else 0
        cost_code_rows.append(
            {
                "cost_code_id": row.cost_code_id,
                "cost_code_description": row.cost_code_description,
                "coins_cost_code": row.coins_cost_code,
                "task_count": task_count,
            }
        )

    html = """
    <!doctype html>
    <html>
    <head>
        <title>Cost Codes</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 40px; background: #f5f6f8; }
            .wrap { max-width: 1100px; margin: 0 auto; }
            .card { background: white; padding: 24px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,.08); margin-bottom: 20px; }
            .btn { padding: 10px 16px; background: #1f4e79; color: white; text-decoration: none; border: none; border-radius: 8px; cursor: pointer; font-weight: bold; margin-right: 8px; }
            table { width: 100%; border-collapse: collapse; }
            th, td { padding: 10px; border-bottom: 1px solid #ddd; text-align: left; vertical-align: top; }
            .top-links a { text-decoration: none; margin-right: 10px; }
            .status-pill { display: inline-block; padding: 4px 8px; border-radius: 999px; background: #e9ecef; }
        </style>
    </head>
    <body>
        <div class="wrap">
            <div class="top-links" style="margin-bottom: 18px;">
                <a href="{{ url_for('home') }}">Home</a>
                <a href="{{ url_for('job_line_items', entity_type=entity_type, entity_id=entity_id) }}">← Back to Line Items</a>
            </div>

            <div class="card">
                <h2>Cost Codes for Line Item {{ line_item_id }}</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Cost Code ID</th>
                            <th>COINS Cost Code</th>
                            <th>Description</th>
                            <th>Task Status</th>
                            <th></th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for row in cost_code_rows %}
                        <tr>
                            <td>{{ row.cost_code_id }}</td>
                            <td>{{ row.coins_cost_code or '' }}</td>
                            <td>{{ row.cost_code_description or '' }}</td>
                            <td>
                                {% if row.task_count and row.task_count > 0 %}
                                    <span class="status-pill">{{ row.task_count }} task(s)</span>
                                {% else %}
                                    <span class="status-pill">Not Scheduled</span>
                                {% endif %}
                            </td>
                            <td>
                                <a class="btn" href="{{ url_for('schedule_cost_code', entity_type=entity_type, entity_id=entity_id, line_item_id=line_item_id, cost_code_id=row.cost_code_id) }}">Schedule Cost Code</a>
                                <a class="btn" href="{{ url_for('edit_cost_code_tasks', entity_type=entity_type, entity_id=entity_id, line_item_id=line_item_id, cost_code_id=row.cost_code_id) }}">Edit Tasks</a>
                            </td>
                        </tr>
                        {% else %}
                        <tr><td colspan="5">No cost codes found.</td></tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </div>
    </body>
    </html>
    """
    return render_template_string(html, cost_code_rows=cost_code_rows, entity_type=entity_type, entity_id=entity_id, line_item_id=line_item_id)
########################################################################################################################
@app.route("/schedule/<entity_type>/<int:entity_id>/line-items/<int:line_item_id>/cost-codes/<int:cost_code_id>", methods=["GET", "POST"])
@office_required
def schedule_cost_code(entity_type, entity_id, line_item_id, cost_code_id):
    templates = get_templates()
    resources = [r for r in get_resources() if int(r.resource_types or 0) != INSTALLATION_RESOURCE_TYPE_ID]
    default_status_id = get_status_id_by_name(DEFAULT_NOT_STARTED_STATUS_NAME)

    if request.method == "POST":
        action = request.form.get("action")

        if action == "load_template":
            template_id = request.form.get("template_id", "").strip()

            if not template_id:
                flash("Please select a template.", "error")
            else:
                template_tasks = fetch_all(
                    """
                    SELECT template_task_id, sequence_number, task_description, resource_type_id
                    FROM task_template_tasks
                    WHERE template_id = ?
                    ORDER BY sequence_number, template_task_id
                    """,
                    [int(template_id)],
                )

                if not template_tasks:
                    flash("Selected template has no tasks.", "error")
                else:
                    conn = get_connection()
                    cur = conn.cursor()
                    for row in template_tasks:
                        cur.execute(
                            """
                            INSERT INTO tasks (cost_code_id, task_name, task_description, sequence_number, resource_type_id, planned_hours, resource_id, status_id)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                            """,
                            cost_code_id,
                            f"Task {row.sequence_number}",
                            row.task_description,
                            row.sequence_number,
                            row.resource_type_id,
                            None,
                            None,
                            default_status_id,
                        )
                    conn.commit()
                    conn.close()
                    flash("Template tasks loaded successfully. Enter hours and optional specific resource constraints on the next screen.", "success")
                    return redirect(url_for('edit_cost_code_tasks', entity_type=entity_type, entity_id=entity_id, line_item_id=line_item_id, cost_code_id=cost_code_id))

        elif action == "create_manual":
            task_name = request.form.get("task_name", "").strip()
            task_description = request.form.get("task_description", "").strip()
            sequence_number = request.form.get("sequence_number", "").strip()
            planned_hours = request.form.get("planned_hours", "").strip()
            resource_type_id = request.form.get("resource_type_id", "").strip()
            resource_id = request.form.get("resource_id", "").strip()

            if not task_name:
                flash("Task name is required.", "error")
            elif not sequence_number:
                flash("Sequence number is required.", "error")
            elif not resource_type_id:
                flash("Resource type is required.", "error")
            else:
                execute_sql(
                    """
                    INSERT INTO tasks (cost_code_id, task_name, task_description, sequence_number, planned_hours, resource_type_id, resource_id, status_id)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    [
                        cost_code_id,
                        task_name,
                        task_description,
                        int(sequence_number),
                        float(planned_hours) if planned_hours else None,
                        int(resource_type_id),
                        int(resource_id) if resource_id else None,
                        default_status_id,
                    ],
                )
                flash("Manual task added successfully.", "success")
                return redirect(url_for('edit_cost_code_tasks', entity_type=entity_type, entity_id=entity_id, line_item_id=line_item_id, cost_code_id=cost_code_id))

    cost_code = fetch_one(
        """
        SELECT cost_code_id, cost_code_description, coins_cost_code, line_item_id
        FROM cost_codes
        WHERE cost_code_id = ?
        """,
        [cost_code_id],
    )

    resource_types = [rt for rt in get_resource_types() if rt.resource_type_id != INSTALLATION_RESOURCE_TYPE_ID]

    html = """
    <!doctype html>
    <html>
    <head>
        <title>Schedule Cost Code</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 40px; background: #f5f6f8; }
            .wrap { max-width: 1100px; margin: 0 auto; }
            .card { background: white; padding: 24px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,.08); margin-bottom: 20px; }
            .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
            input[type='text'], input[type='number'], textarea, select {
                width: 100%; padding: 10px; border: 1px solid #ccc; border-radius: 6px;
                margin-top: 6px; margin-bottom: 14px; box-sizing: border-box;
            }
            .btn { padding: 10px 16px; background: #1f4e79; color: white; text-decoration: none; border: none; border-radius: 8px; cursor: pointer; font-weight: bold; }
            .top-links a { text-decoration: none; margin-right: 10px; }
            .flash-success { background: #d1e7dd; color: #0f5132; padding: 10px; border-radius: 6px; margin-bottom: 12px; }
            .flash-error { background: #f8d7da; color: #842029; padding: 10px; border-radius: 6px; margin-bottom: 12px; }
        </style>
    </head>
    <body>
        <div class="wrap">
            <div class="top-links" style="margin-bottom: 18px;">
                <a href="{{ url_for('home') }}">Home</a>
                <a href="{{ url_for('line_item_cost_codes', entity_type=entity_type, entity_id=entity_id, line_item_id=line_item_id) }}">← Back to Cost Codes</a>
            </div>

            {% with messages = get_flashed_messages(with_categories=true) %}
                {% if messages %}
                    {% for category, message in messages %}
                        <div class="flash-{{ category }}">{{ message }}</div>
                    {% endfor %}
                {% endif %}
            {% endwith %}

            <div class="card">
                <h2>Schedule Cost Code</h2>
                <p><strong>Cost Code ID:</strong> {{ cost_code.cost_code_id }}</p>
                <p><strong>COINS Cost Code:</strong> {{ cost_code.coins_cost_code or '' }}</p>
                <p><strong>Description:</strong> {{ cost_code.cost_code_description or '' }}</p>
                <p><strong>Note:</strong> The scheduler defines sequence, planned hours, and resource type. A specific resource is only an optional hard constraint.</p>
            </div>

            <div class="grid">
                <div class="card">
                    <h3>Load Template</h3>
                    <form method="post">
                        <input type="hidden" name="action" value="load_template">

                        <label>Template</label>
                        <select name="template_id" required>
                            <option value="">-- Select Template --</option>
                            {% for t in templates %}
                                <option value="{{ t.template_id }}">{{ t.template_name }}</option>
                            {% endfor %}
                        </select>

                        <p style="margin-top: 10px; color: #555;">Loading a template copies the task structure and resource types only. You will enter planned hours and any optional specific resource constraints on the next screen.</p>

                        <button class="btn" type="submit">Load Template</button>
                    </form>
                </div>

                <div class="card">
                    <h3>Create Manual Task</h3>
                    <form method="post">
                        <input type="hidden" name="action" value="create_manual">

                        <label>Sequence Number</label>
                        <input type="number" name="sequence_number" required>

                        <label>Task Name</label>
                        <input type="text" name="task_name" required>

                        <label>Task Description</label>
                        <textarea name="task_description" rows="3"></textarea>

                        <label>Planned Hours</label>
                        <input type="number" step="0.1" name="planned_hours[]">

                        <label>Resource Type</label>
                        <select name="resource_type_id" class="resource-type-select" required>
                            <option value="">-- Select Resource Type --</option>
                            {% for rt in resource_types %}
                                <option value="{{ rt.resource_type_id }}">{{ rt.resource_description }}</option>
                            {% endfor %}
                        </select>

                        <label>Specific Resource (Optional Constraint)</label>
                        <select name="resource_id" class="resource-select">
                            <option value="">-- Optional --</option>
                            {% for r in resources %}
                                <option value="{{ r.resource_id }}" data-resource-type="{{ r.resource_types }}">
                                    {{ r.resource_name }}
                                </option>
                            {% endfor %}
                        </select>

                        <button class="btn" type="submit">Add Manual Task</button>
                    </form>
                </div>
            </div>
        </div>
    </body>
    </html>
    """
    return render_template_string(
        html,
        cost_code=cost_code,
        templates=templates,
        resources=resources,
        resource_types=resource_types,
        entity_type=entity_type,
        entity_id=entity_id,
        line_item_id=line_item_id,
    )
########################################################################################################################
@app.route("/schedule/<entity_type>/<int:entity_id>/line-items/<int:line_item_id>/cost-codes/<int:cost_code_id>/tasks", methods=["GET", "POST"])
@office_required
def edit_cost_code_tasks(entity_type, entity_id, line_item_id, cost_code_id):
    resources = [r for r in get_resources() if int(r.resource_types or 0) != INSTALLATION_RESOURCE_TYPE_ID]
    resource_types = [rt for rt in get_resource_types() if rt.resource_type_id != INSTALLATION_RESOURCE_TYPE_ID]

    if request.method == "POST":
        action = request.form.get("action")

        if action == "bulk_save_hours":
            task_ids = request.form.getlist("task_id[]")
            planned_hours_list = request.form.getlist("planned_hours[]")

            conn = get_connection()
            cur = conn.cursor()
            for i, task_id in enumerate(task_ids):
                planned_hours = planned_hours_list[i].strip() if i < len(planned_hours_list) and planned_hours_list[i] is not None else ""
                cur.execute(
                    """
                    UPDATE tasks
                    SET planned_hours = ?
                    WHERE task_id = ?
                    """,
                    float(planned_hours) if planned_hours else None,
                    int(task_id),
                )
            conn.commit()
            conn.close()
            flash("Hours saved for all tasks.", "success")
            return redirect(url_for(
                "edit_cost_code_tasks",
                entity_type=entity_type,
                entity_id=entity_id,
                line_item_id=line_item_id,
                cost_code_id=cost_code_id,
            ))

        elif action == "update_task":
            task_id = int(request.form.get("task_id"))
            task_name = request.form.get("task_name", "").strip()
            task_description = request.form.get("task_description", "").strip()
            sequence_number = request.form.get("sequence_number", "").strip()
            planned_hours = request.form.get("planned_hours_single", "").strip()
            resource_type_id = request.form.get("resource_type_id", "").strip()
            resource_id = request.form.get("resource_id", "").strip()

            if not task_name:
                flash("Task name is required.", "error")
            elif not sequence_number:
                flash("Sequence number is required.", "error")
            elif not resource_type_id:
                flash("Resource type is required.", "error")
            else:
                execute_sql(
                    """
                    UPDATE tasks
                    SET task_name = ?,
                        task_description = ?,
                        sequence_number = ?,
                        planned_hours = ?,
                        resource_type_id = ?,
                        resource_id = ?
                    WHERE task_id = ?
                    """,
                    [
                        task_name,
                        task_description,
                        int(sequence_number),
                        float(planned_hours) if planned_hours else None,
                        int(resource_type_id),
                        int(resource_id) if resource_id else None,
                        task_id,
                    ],
                )
                flash("Task updated successfully.", "success")
                return redirect(url_for(
                    "edit_cost_code_tasks",
                    entity_type=entity_type,
                    entity_id=entity_id,
                    line_item_id=line_item_id,
                    cost_code_id=cost_code_id,
                ))

        elif action == "delete_task":
            task_id = int(request.form.get("task_id"))
            execute_sql("DELETE FROM tasks WHERE task_id = ?", [task_id])
            flash("Task deleted successfully.", "success")
            return redirect(url_for(
                "edit_cost_code_tasks",
                entity_type=entity_type,
                entity_id=entity_id,
                line_item_id=line_item_id,
                cost_code_id=cost_code_id,
            ))

    cost_code = fetch_one(
        """
        SELECT cost_code_id, cost_code_description, coins_cost_code
        FROM cost_codes
        WHERE cost_code_id = ?
        """,
        [cost_code_id],
    )

    line_item = fetch_one(
        """
        SELECT line_item_id, quantity, line_item_description
        FROM line_items
        WHERE line_item_id = ?
        """,
        [line_item_id],
    )

    tasks = fetch_all(
        """
        SELECT
            t.task_id,
            t.task_name,
            t.task_description,
            t.sequence_number,
            t.planned_hours,
            t.resource_type_id,
            t.resource_id,
            t.status_id
        FROM tasks AS t
        WHERE t.cost_code_id = ?
        ORDER BY t.sequence_number, t.task_id
        """,
        [cost_code_id],
    )

    resource_name_map = get_resource_name_map()
    resource_type_name_map = get_resource_type_name_map()
    status_name_map = get_status_name_map()

    task_rows = []
    for row in tasks:
        task_rows.append(
            {
                "task_id": row.task_id,
                "task_name": row.task_name,
                "task_description": row.task_description,
                "sequence_number": row.sequence_number,
                "planned_hours": row.planned_hours,
                "resource_type_id": row.resource_type_id,
                "resource_id": row.resource_id,
                "resource_name": resource_name_map.get(row.resource_id, "") if row.resource_id is not None else "",
                "resource_description": resource_type_name_map.get(row.resource_type_id, "") if row.resource_type_id is not None else "",
                "status_names": status_name_map.get(row.status_id, "") if row.status_id is not None else "",
            }
        )

    html = """
    <!doctype html>
    <html>
    <head>
    <title>Edit Tasks</title>
    <script>
        function filterResourceOptions(row) {
            const typeSelect = row.querySelector('.resource-type-select');
            const resourceSelect = row.querySelector('.resource-select');

            if (!typeSelect || !resourceSelect) return;

            const selectedType = String(typeSelect.value || '');
            const currentValue = String(resourceSelect.value || '');

            Array.from(resourceSelect.options).forEach((opt, idx) => {
                if (idx === 0) {
                    opt.hidden = false; // keep the blank optional row visible
                    return;
                }

                const optType = String(opt.getAttribute('data-resource-type') || '');

                if (selectedType === '') {
                    opt.hidden = false;
                } else {
                    opt.hidden = (optType !== selectedType);
                }
            });

            if (currentValue) {
                const selectedOption = Array.from(resourceSelect.options).find(
                    o => o.value === currentValue
                );

                const selectedOptionType = selectedOption
                    ? String(selectedOption.getAttribute('data-resource-type') || '')
                    : '';

                if (selectedType && selectedOptionType !== selectedType) {
                    resourceSelect.value = '';
                }
            }
        }

        function initializeResourceFilters() {
            document.querySelectorAll('tbody tr').forEach((row) => {
                const typeSelect = row.querySelector('.resource-type-select');
                if (!typeSelect) return;

                filterResourceOptions(row);

                typeSelect.addEventListener('change', function () {
                    filterResourceOptions(row);
                });
            });
        }

        document.addEventListener('DOMContentLoaded', initializeResourceFilters);
    </script>

    <style>
        body { font-family: Arial, sans-serif; margin: 40px; background: #f5f6f8; }
        .wrap { max-width: 1500px; margin: 0 auto; }
        .card { background: white; padding: 24px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,.08); margin-bottom: 20px; }
        table { width: 100%; border-collapse: collapse; }
        th, td { padding: 10px; border-bottom: 1px solid #ddd; text-align: left; vertical-align: top; }
        input[type='text'], input[type='number'], textarea, select {
            width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 6px; box-sizing: border-box;
        }
        .btn { padding: 8px 12px; background: #1f4e79; color: white; text-decoration: none; border: none; border-radius: 8px; cursor: pointer; font-weight: bold; margin-bottom: 4px; }
        .btn-delete { background: #9b2226; }
        .top-links a { text-decoration: none; margin-right: 10px; }
        .flash-success { background: #d1e7dd; color: #0f5132; padding: 10px; border-radius: 6px; margin-bottom: 12px; }
        .flash-error { background: #f8d7da; color: #842029; padding: 10px; border-radius: 6px; margin-bottom: 12px; }
        .actions { display: flex; flex-direction: column; gap: 6px; }
        .desc-box { min-width: 220px; }
    </style>
</head>
<body>
    <div class="wrap">
        <div class="top-links" style="margin-bottom: 18px;">
            <a href="{{ url_for('home') }}">Home</a>
            <a href="{{ url_for('line_item_cost_codes', entity_type=entity_type, entity_id=entity_id, line_item_id=line_item_id) }}">← Back to Cost Codes</a>
        </div>

        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}
                {% for category, message in messages %}
                    <div class="flash-{{ category }}">{{ message }}</div>
                {% endfor %}
            {% endif %}
        {% endwith %}

        <div class="card">
            <h2>Edit Tasks for Cost Code {{ cost_code.cost_code_id }}</h2>
            <p><strong>COINS Cost Code:</strong> {{ cost_code.coins_cost_code or '' }}</p>
            <p><strong>Description:</strong> {{ cost_code.cost_code_description or '' }}</p>
            <p style="font-size: 28px; font-weight: bold; margin: 10px 0;"><strong>Qty:</strong> {{ line_item.quantity or '' }}</p>
            <p><strong>Line Item:</strong> {{ line_item.line_item_description or '' }}</p>
            <p><a class="btn" href="{{ url_for('schedule_cost_code', entity_type=entity_type, entity_id=entity_id, line_item_id=line_item_id, cost_code_id=cost_code_id) }}">Add More Tasks / Load Template</a></p>
        </div>

        <div class="card">
            <form method="post" style="margin-bottom: 16px;">
                {% for row in task_rows %}
                    <input type="hidden" name="task_id[]" value="{{ row.task_id }}">
                {% endfor %}

                <table>
                    <thead>
                        <tr>
                            <th>Task ID</th>
                            <th>Sequence</th>
                            <th>Task Name</th>
                            <th>Description</th>
                            <th>Hours</th>
                            <th>Resource Type</th>
                            <th>Specific Resource (Optional)</th>
                            <th>Status</th>
                            <th>Actions</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for row in task_rows %}
                        <tr>
                            <td>{{ row.task_id }}</td>
                            <td>
                                <input type="number" name="sequence_number_display" value="{{ row.sequence_number or '' }}" disabled>
                            </td>
                            <td>
                                <input type="text" name="task_name_display" value="{{ row.task_name or '' }}" disabled>
                            </td>
                            <td class="desc-box">
                                <textarea name="task_description_display" rows="2" disabled>{{ row.task_description or '' }}</textarea>
                            </td>
                            <td>
                                <input type="number" step="0.1" name="planned_hours[]" value="{{ row.planned_hours or '' }}">
                            </td>
                            <td>
                                <form method="post">
                                    <input type="hidden" name="task_id" value="{{ row.task_id }}">
                                    <input type="hidden" name="sequence_number" value="{{ row.sequence_number or '' }}">
                                    <input type="hidden" name="task_name" value="{{ row.task_name or '' }}">
                                    <input type="hidden" name="task_description" value="{{ row.task_description or '' }}">
                                    <input type="hidden" name="planned_hours_single" value="{{ row.planned_hours or '' }}">

                                    <select name="resource_type_id" class="resource-type-select" required>
                                        <option value="">-- Select Resource Type --</option>
                                        {% for rt in resource_types %}
                                            <option value="{{ rt.resource_type_id }}" {% if row.resource_type_id == rt.resource_type_id %}selected{% endif %}>
                                                {{ rt.resource_description }}
                                            </option>
                                        {% endfor %}
                                    </select>
                            </td>
                            <td>
                                    <select name="resource_id" class="resource-select">
                                        <option value="">-- Optional --</option>
                                        {% for r in resources %}
                                            <option value="{{ r.resource_id }}"
                                                    data-resource-type="{{ r.resource_types }}"
                                                    {% if row.resource_id == r.resource_id %}selected{% endif %}>
                                                {{ r.resource_name }}
                                            </option>
                                        {% endfor %}
                                    </select>
                            </td>
                            <td>{{ row.status_names or '' }}</td>
                            <td>
                                <div class="actions">
                                    <button class="btn" type="submit" name="action" value="update_task">Save</button>
                                    <button class="btn btn-delete" type="submit" name="action" value="delete_task" onclick="return confirm('Delete this task?');">Delete</button>
                                </div>
                                </form>
                            </td>
                        </tr>
                        {% else %}
                        <tr><td colspan="9">No tasks exist for this cost code yet.</td></tr>
                        {% endfor %}
                    </tbody>
                </table>

                <div style="margin-top: 16px;">
                    <button class="btn" type="submit" name="action" value="bulk_save_hours">Save All Hours</button>
                </div>
            </form>
        </div>
    </div>
</body>
    </html>
    """
    return render_template_string(
        html,
        cost_code=cost_code,
        line_item=line_item,
        task_rows=task_rows,
        resources=resources,
        resource_types=resource_types,
        entity_type=entity_type,
        entity_id=entity_id,
        line_item_id=line_item_id,
        cost_code_id=cost_code_id,
    )
########################################################################################################################
########################################################################################################################
########################################################################################################################





########################################################################################################################
########################################################################################################################
########################################################################################################################
"""INSTALLATION SCHEDULING"""
########################################################################################################################
########################################################################################################################
########################################################################################################################
@app.route("/installation")
@office_required
def installation_calendar():
    from calendar import monthrange

    year = request.args.get("year", default=datetime.now().year, type=int)
    month = request.args.get("month", default=datetime.now().month, type=int)

    if month < 1:
        month = 12
        year -= 1
    elif month > 12:
        month = 1
        year += 1

    _, days_in_month = monthrange(year, month)
    first_weekday, _ = monthrange(year, month)
    sun_offset = (first_weekday + 1) % 7

    month_start = date(year, month, 1)
    month_end = date(year, month, days_in_month)

    prev_month = month - 1 if month > 1 else 12
    prev_year = year if month > 1 else year - 1
    next_month = month + 1 if month < 12 else 1
    next_year = year if month < 12 else year + 1

    month_name = date(year, month, 1).strftime("%B %Y")

    calendar_tasks = fetch_all(
        """
        SELECT
            t.task_id,
            t.task_name,
            t.planned_start_date,
            t.planned_finish_date,
            t.planned_hours,
            r.resource_name
        FROM tasks AS t
        LEFT JOIN resources AS r
            ON t.resource_id = r.resource_id
        WHERE t.resource_type_id = ?
          AND t.planned_start_date IS NOT NULL
          AND t.planned_finish_date IS NOT NULL
        ORDER BY t.planned_start_date, t.task_id
        """,
        [INSTALLATION_RESOURCE_TYPE_ID],
    )

    day_map = {}

    for task in calendar_tasks:
        task_day = (
            task.planned_start_date.date()
            if isinstance(task.planned_start_date, datetime)
            else task.planned_start_date
        )

        if not task_day or not (month_start <= task_day <= month_end):
            continue

        day_key = task_day.strftime("%Y-%m-%d")

        day_map.setdefault(day_key, []).append(
            {
                "task_name": task.task_name or f"Task {task.task_id}",
                "resource_name": task.resource_name or "Unassigned",
                "hours": round(float(task.planned_hours or 0), 1),
            }
        )

    jobs = fetch_all(
        """
        SELECT
            j.job_id,
            j.due_date,
            j.manufacturing_due_date,
            j.manufacturing_scheduled,
            j.survey_scheduled,
            j.installation_scheduled,
            c.customer_name
        FROM jobs AS j
        LEFT JOIN customers AS c
            ON j.customer_id = c.customer_id
        ORDER BY
            j.manufacturing_due_date,
            j.due_date,
            j.job_id
        """
    )

    survey_not_scheduled = []
    survey_scheduled = []
    installation_scheduled = []

    for job in jobs:
        row = {
            "job_id": job.job_id,
            "customer": job.customer_name or "",
            "manufacturing_due_date": job.manufacturing_due_date,
            "due_date": job.due_date,
        }

        if is_yes_no_true(job.installation_scheduled):
            installation_scheduled.append(row)
        elif is_yes_no_true(job.survey_scheduled):
            survey_scheduled.append(row)
        else:
            survey_not_scheduled.append(row)

    today = datetime.now().date()

    html = """
    <!doctype html>
    <html>
    <head>
        <title>Installation Scheduling</title>
        <style>
            * { box-sizing:border-box; }
            body { font-family:Arial,sans-serif; margin:0; padding:22px; background:#f5f6f8; }
            .wrap { max-width:1700px; margin:auto; }
            .card { background:white; padding:20px; border-radius:10px; box-shadow:0 2px 10px rgba(0,0,0,.08); margin-bottom:18px; }
            .btn { display:inline-block; padding:10px 18px; background:#1f4e79; color:white; text-decoration:none; border-radius:7px; font-weight:bold; }
            .btn-back { background:#6c757d; margin-bottom:16px; }
            .topbar { display:flex; justify-content:space-between; align-items:center; }
            .month-title { font-size:26px; font-weight:bold; color:#1f4e79; }

            .cal-grid { display:grid; grid-template-columns:repeat(7,1fr); gap:4px; }
            .cal-dow { text-align:center; font-weight:bold; padding:8px; background:#1f4e79; color:white; border-radius:6px; }
            .cal-day { min-height:110px; border:1px solid #ddd; border-radius:6px; padding:6px; background:white; }
            .cal-day.today { border:2px solid #1f4e79; }
            .cal-day.empty { background:#f8f9fa; border-color:#eee; }
            .day-num { font-weight:bold; margin-bottom:4px; }
            .chip { background:#198754; color:white; font-size:10px; border-radius:4px; padding:3px 5px; margin-bottom:3px; overflow-wrap:anywhere; }

            .pipeline { display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:16px; }
            .column { background:white; border-radius:10px; box-shadow:0 2px 10px rgba(0,0,0,.08); overflow:hidden; }
            .column-head { color:white; font-size:17px; font-weight:bold; padding:15px; display:flex; justify-content:space-between; }
            .orange { background:#fd7e14; }
            .blue { background:#0d6efd; }
            .green { background:#198754; }
            .column-body { padding:14px; max-height:450px; overflow:auto; }
            .search { width:100%; padding:9px; border:1px solid #ccc; border-radius:7px; margin-bottom:10px; }
            .job-card { border:1px solid #ddd; border-radius:8px; padding:11px; margin-bottom:9px; cursor:pointer; background:white; }
            .job-card:hover { background:#eef4fb; }
            .job-card.s1 { border-left:5px solid #fd7e14; }
            .job-card.s2 { border-left:5px solid #0d6efd; }
            .job-card.s3 { border-left:5px solid #198754; }
            .muted { color:#777; }

            @media(max-width:1000px){ .pipeline { grid-template-columns:1fr; } }
        </style>
        <script>
            function filterCards(inputId, cls) {
                const q = document.getElementById(inputId).value.toLowerCase();
                document.querySelectorAll("." + cls).forEach(function(card) {
                    card.style.display = card.innerText.toLowerCase().includes(q) ? "block" : "none";
                });
            }
        </script>
    </head>
    <body>
    <div class="wrap">

        <a class="btn btn-back" href="{{ url_for('home') }}">← Home</a>

        <div class="card">
            <div class="topbar">
                <a class="btn" href="{{ url_for('installation_calendar', year=prev_year, month=prev_month) }}">‹ Prev</a>
                <div class="month-title">{{ month_name }}</div>
                <a class="btn" href="{{ url_for('installation_calendar', year=next_year, month=next_month) }}">Next ›</a>
            </div>
        </div>

        <div class="card">
            <div class="cal-grid">
                {% for dow in ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'] %}
                    <div class="cal-dow">{{ dow }}</div>
                {% endfor %}

                {% for _ in range(sun_offset) %}
                    <div class="cal-day empty"></div>
                {% endfor %}

                {% for d in range(1, days_in_month + 1) %}
                    {% set day_date = "%04d-%02d-%02d"|format(year, month, d) %}
                    {% set chips = day_map.get(day_date, []) %}

                    <div class="cal-day {% if d == today.day and month == today.month and year == today.year %}today{% endif %}">
                        <div class="day-num">{{ d }}</div>

                        {% for chip in chips[:5] %}
                            <div class="chip">
                                {{ chip.task_name }}<br>
                                {{ chip.resource_name }} — {{ chip.hours }}h
                            </div>
                        {% endfor %}

                        {% if chips|length > 5 %}
                            <small>+{{ chips|length - 5 }} more</small>
                        {% endif %}
                    </div>
                {% endfor %}
            </div>
        </div>

        <div class="pipeline">

            <div class="column">
                <div class="column-head orange">
                    <span>Survey Not Scheduled</span>
                    <span>{{ survey_not_scheduled|length }}</span>
                </div>
                <div class="column-body">
                    <input id="q1" class="search" placeholder="Search..." onkeyup="filterCards('q1','c1')">
                    {% for row in survey_not_scheduled %}
                        <div class="job-card s1 c1"
                             onclick="window.location='{{ url_for('installation_job', job_id=row.job_id, stage='survey') }}'">
                            <strong>Job {{ row.job_id }}</strong><br>
                            {{ row.customer }}<br>
                            <small>
                            Mfg Due:
                            {{ row.manufacturing_due_date.strftime('%b %d, %Y') if row.manufacturing_due_date else '—' }}
                            </small>

                            <br>
                            
                            <small>
                                Project Completion:
                                {{ row.due_date.strftime('%b %d, %Y') if row.due_date else '—' }}
                            </small>
                        </div>
                    {% else %}
                        <p class="muted">No jobs waiting for survey scheduling.</p>
                    {% endfor %}
                </div>
            </div>

            <div class="column">
                <div class="column-head blue">
                    <span>Survey Scheduled / Install Not Scheduled</span>
                    <span>{{ survey_scheduled|length }}</span>
                </div>
                <div class="column-body">
                    <input id="q2" class="search" placeholder="Search..." onkeyup="filterCards('q2','c2')">
                    {% for row in survey_scheduled %}
                        <div class="job-card s2 c2"
                             onclick="window.location='{{ url_for('installation_job', job_id=row.job_id, stage='install') }}'">
                            <strong>Job {{ row.job_id }}</strong><br>
                            {{ row.customer }}<br>
                            <small>
                            Mfg Due:
                            {{ row.manufacturing_due_date.strftime('%b %d, %Y') if row.manufacturing_due_date else '—' }}
                            </small>

                            <br>
                            
                            <small>
                                Project Completion:
                                {{ row.due_date.strftime('%b %d, %Y') if row.due_date else '—' }}
                            </small>
                        </div>
                    {% else %}
                        <p class="muted">No jobs waiting for installation scheduling.</p>
                    {% endfor %}
                </div>
            </div>

            <div class="column">
                <div class="column-head green">
                    <span>Installation Scheduled</span>
                    <span>{{ installation_scheduled|length }}</span>
                </div>
                <div class="column-body">
                    <input id="q3" class="search" placeholder="Search..." onkeyup="filterCards('q3','c3')">
                    {% for row in installation_scheduled %}
                        <div class="job-card s3 c3"
                             onclick="window.location='{{ url_for('installation_job', job_id=row.job_id, stage='install') }}'">
                            <strong>Job {{ row.job_id }}</strong><br>
                            {{ row.customer }}<br>
                            <small>
                            Mfg Due:
                            {{ row.manufacturing_due_date.strftime('%b %d, %Y') if row.manufacturing_due_date else '—' }}
                            </small>

                            <br>
                            
                            <small>
                                Project Completion:
                                {{ row.due_date.strftime('%b %d, %Y') if row.due_date else '—' }}
                            </small>
                        </div>
                    {% else %}
                        <p class="muted">No scheduled installations.</p>
                    {% endfor %}
                </div>
            </div>

        </div>
    </div>
    </body>
    </html>
    """

    return render_template_string(
        html,
        month_name=month_name,
        year=year,
        month=month,
        days_in_month=days_in_month,
        sun_offset=sun_offset,
        day_map=day_map,
        prev_year=prev_year,
        prev_month=prev_month,
        next_year=next_year,
        next_month=next_month,
        today=today,
        survey_not_scheduled=survey_not_scheduled,
        survey_scheduled=survey_scheduled,
        installation_scheduled=installation_scheduled,
    )
########################################################################################################################
@app.route("/installation/job/<int:job_id>", methods=["GET", "POST"])
@office_required
def installation_job(job_id):
    from calendar import monthrange

    header = fetch_one(
        """
        SELECT
            j.job_id,
            j.due_date,
            j.manufacturing_due_date,
            j.manufacturing_scheduled,
            j.survey_scheduled,
            j.installation_scheduled,
            c.customer_name
        FROM jobs AS j
        LEFT JOIN customers AS c
            ON j.customer_id = c.customer_id
        WHERE j.job_id = ?
        """,
        [job_id],
    )

    if not header:
        flash("Job not found.", "error")
        return redirect(url_for("installation_calendar"))

    first_cc = _first_job_cost_code(job_id)

    if not first_cc:
        flash(
            "No cost codes found for this job. Add cost codes first.",
            "error",
        )
        return redirect(url_for("installation_calendar"))

    stage = (
        request.form.get("stage", "").strip()
        if request.method == "POST"
        else request.args.get("stage", "").strip()
    )

    if stage not in ("survey", "install"):
        stage = (
            "survey"
            if not is_yes_no_true(header.survey_scheduled)
            else "install"
        )

    year = request.args.get(
        "year",
        default=datetime.now().year,
        type=int,
    )

    month = request.args.get(
        "month",
        default=datetime.now().month,
        type=int,
    )

    if month < 1:
        month = 12
        year -= 1

    elif month > 12:
        month = 1
        year += 1

    _, days_in_month = monthrange(year, month)

    first_weekday, _ = monthrange(year, month)
    sun_offset = (first_weekday + 1) % 7

    month_start = date(year, month, 1)
    month_end = date(year, month, days_in_month)

    prev_month = month - 1 if month > 1 else 12
    prev_year = year if month > 1 else year - 1

    next_month = month + 1 if month < 12 else 1
    next_year = year if month < 12 else year + 1

    month_name = date(
        year,
        month,
        1
    ).strftime("%B %Y")

    resources = _installation_resources()

    survey_cost_codes = _survey_cost_code_options(
        job_id
    )

    survey_cost_code_options = [
        {
            "cost_code_id": int(row.cost_code_id),
            "coins_cost_code": row.coins_cost_code or "",
            "cost_code_description": row.cost_code_description or "",
            "line_item_description": row.line_item_description or "",
        }
        for row in survey_cost_codes
    ]

    qc_forms = _installation_qc_forms()

    qc_form_options = [
        {
            "qc_form_id": int(row.qc_form_id),
            "form_name": (
                    row.form_name
                    or f"Form {row.qc_form_id}"
            ),
        }
        for row in qc_forms
    ]

    valid_qc_ids = {
        row["qc_form_id"]
        for row in qc_form_options
    }

    existing_tasks = _job_rt13_tasks(job_id)

    # =============================================================================
    # CREATE SURVEY TASKS
    # =============================================================================

    if request.method == "POST" and stage == "survey":

        selected_dates_raw = request.form.get(
            "selected_dates",
            "",
        ).strip()

        resource_id_values = request.form.getlist(
            "survey_resource_ids"
        )

        resource_ids = []

        for value in resource_id_values:
            try:
                resource_ids.append(int(value))
            except (TypeError, ValueError):
                pass

        survey_count_text = request.form.get(
            "survey_count",
            "",
        ).strip()

        try:
            survey_count = int(survey_count_text)

        except (TypeError, ValueError):
            survey_count = 0

        selected_dates = [
            value.strip()
            for value in selected_dates_raw.split(",")
            if value.strip()
        ]

        # -------------------------------------------------------------------------
        # Validation
        # -------------------------------------------------------------------------

        if not selected_dates:
            flash(
                "Select at least one survey date.",
                "error",
            )

            return redirect(
                url_for(
                    "installation_job",
                    job_id=job_id,
                    stage="survey",
                    year=year,
                    month=month,
                )
            )

        if not resource_ids:
            flash(
                "Select at least one survey team.",
                "error",
            )

            return redirect(
                url_for(
                    "installation_job",
                    job_id=job_id,
                    stage="survey",
                    year=year,
                    month=month,
                )
            )

        if survey_count <= 0:
            flash(
                "Enter the number of surveys.",
                "error",
            )

            return redirect(
                url_for(
                    "installation_job",
                    job_id=job_id,
                    stage="survey",
                    year=year,
                    month=month,
                )
            )

        # -------------------------------------------------------------------------
        # Get the QC form selected for each survey
        # -------------------------------------------------------------------------

        selected_form_ids = []

        for i in range(1, survey_count + 1):

            try:
                form_id = int(
                    request.form.get(
                        f"survey_form_{i}",
                        "",
                    )
                )

            except (TypeError, ValueError):
                form_id = 0

            if form_id not in valid_qc_ids:
                flash(
                    f"Select a survey type for Survey {i}.",
                    "error",
                )

                return redirect(
                    url_for(
                        "installation_job",
                        job_id=job_id,
                        stage="survey",
                        year=year,
                        month=month,
                    )
                )

            selected_form_ids.append(form_id)

        form_name_map = {
            row["qc_form_id"]: row["form_name"]
            for row in qc_form_options
        }

        selected_date_objects = [
            datetime.strptime(
                value,
                "%Y-%m-%d",
            ).date()
            for value in selected_dates
        ]

        selected_date_objects.sort()

        # All survey tasks share the same overall survey window.
        survey_start_date = selected_date_objects[0]
        survey_finish_date = selected_date_objects[-1]

        start_dt = datetime.combine(
            survey_start_date,
            time(7, 0),
        )

        finish_dt = datetime.combine(
            survey_finish_date,
            time(16, 0),
        )

        conn = get_connection()
        cur = conn.cursor()

        for i, form_id in enumerate(
                selected_form_ids,
                start=1,
        ):
            # All survey tasks are assigned to the first selected installation team.
            # Any installer in the department can still see them through All Department Tasks.
            resource_id = resource_ids[0]

            form_name = form_name_map.get(
                form_id,
                f"Survey {i}",
            )

            survey_suffix = request.form.get(
                f"survey_suffix_{i}",
                "",
            ).strip()

            try:
                selected_cost_code_id = int(
                    request.form.get(
                        f"survey_cost_code_{i}",
                        "",
                    )
                )

            except (TypeError, ValueError):
                selected_cost_code_id = 0

            if not selected_cost_code_id:
                flash(
                    f"Select a cost code for Survey {i}.",
                    "error",
                )

                conn.close()

                return redirect(
                    url_for(
                        "installation_job",
                        job_id=job_id,
                        stage="survey",
                        year=year,
                        month=month,
                    )
                )

            valid_cost_code = fetch_one(
                """
                SELECT
                    cc.cost_code_id
                FROM line_items AS li
                INNER JOIN cost_codes AS cc
                    ON cc.line_item_id = li.line_item_id
                WHERE li.job_number_id = ?
                  AND cc.cost_code_id = ?
                """,
                [
                    job_id,
                    selected_cost_code_id,
                ],
            )

            if not valid_cost_code:
                flash(
                    f"The selected cost code for Survey {i} does not belong to this job.",
                    "error",
                )

                conn.close()

                return redirect(
                    url_for(
                        "installation_job",
                        job_id=job_id,
                        stage="survey",
                        year=year,
                        month=month,
                    )
                )

            if not survey_suffix:
                flash(
                    f"Enter task text for Survey {i}.",
                    "error",
                )
                conn.close()
                return redirect(
                    url_for(
                        "installation_job",
                        job_id=job_id,
                        stage="survey",
                        year=year,
                        month=month,
                    )
                )

            cur.execute(
                """
                INSERT INTO tasks
                    (
                        cost_code_id,
                        task_name,
                        task_description,
                        sequence_number,
                        resource_type_id,
                        resource_id,
                        planned_hours,
                        planned_start_date,
                        planned_finish_date,
                        status_id,
                        survey_type
                    )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                selected_cost_code_id,
                f"Survey — {form_name} — {survey_suffix}",
                f"Survey {i} of {survey_count}",
                90 + i,
                INSTALLATION_RESOURCE_TYPE_ID,
                resource_id,
                None,
                start_dt,
                finish_dt,
                1,
                form_id,
            )

        cur.execute(
            """
            UPDATE jobs
            SET survey_scheduled = True
            WHERE job_id = ?
            """,
            job_id,
        )

        conn.commit()
        conn.close()

        flash(
            f"Created {survey_count} individual survey task(s) "
            f"across {len(selected_date_objects)} selected date(s).",
            "success",
        )

        return redirect(
            url_for(
                "installation_job",
                job_id=job_id,
                stage="install",
                year=year,
                month=month,
            )
        )

    # =============================================================================
    # CREATE INSTALLATION TASKS
    # =============================================================================

    if request.method == "POST" and stage == "install":

        selected_dates_raw = request.form.get(
            "selected_dates",
            "",
        ).strip()

        if not selected_dates_raw:
            flash(
                "Select at least one installation date.",
                "error",
            )

            return redirect(
                url_for(
                    "installation_job",
                    job_id=job_id,
                    stage="install",
                    year=year,
                    month=month,
                )
            )

        selected_dates = [
            x.strip()
            for x in selected_dates_raw.split(",")
            if x.strip()
        ]

        assignments_by_resource = {}

        for selected_date in selected_dates:

            for resource in resources:

                hours_text = request.form.get(
                    f"hours_{selected_date}_{resource.resource_id}",
                    "",
                ).strip()

                if not hours_text:
                    continue

                try:
                    hours_value = float(hours_text)

                except ValueError:
                    continue

                if hours_value <= 0:
                    continue

                assignments_by_resource.setdefault(
                    resource.resource_id,
                    {
                        "dates": [],
                        "total_hours": 0.0,
                    },
                )

                assignments_by_resource[
                    resource.resource_id
                ]["dates"].append(
                    {
                        "date": selected_date,
                        "hours": hours_value,
                    }
                )

                assignments_by_resource[
                    resource.resource_id
                ]["total_hours"] += hours_value

        if not assignments_by_resource:
            flash(
                "No installation hours were entered.",
                "error",
            )

            return redirect(
                url_for(
                    "installation_job",
                    job_id=job_id,
                    stage="install",
                    year=year,
                    month=month,
                )
            )

        install_form_id = (
            _installation_completion_form_id()
        )

        conn = get_connection()
        cur = conn.cursor()

        for resource_id, assignment in (
                assignments_by_resource.items()
        ):
            dates = sorted(
                assignment["dates"],
                key=lambda x: x["date"],
            )

            first_date = dates[0]["date"]
            last_date = dates[-1]["date"]
            last_day_hours = dates[-1]["hours"]

            start_dt = datetime.strptime(
                f"{first_date} 07:00",
                "%Y-%m-%d %H:%M",
            )

            finish_start = datetime.strptime(
                f"{last_date} 07:00",
                "%Y-%m-%d %H:%M",
            )

            finish_dt = (
                    finish_start
                    + timedelta(hours=last_day_hours)
            )

            cur.execute(
                """
                INSERT INTO tasks
                    (
                        cost_code_id,
                        task_name,
                        task_description,
                        sequence_number,
                        resource_type_id,
                        resource_id,
                        planned_hours,
                        planned_start_date,
                        planned_finish_date,
                        status_id,
                        survey_type
                    )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                first_cc.cost_code_id,
                f"Installation — Job {job_id}",
                "Scheduled installation",
                99,
                INSTALLATION_RESOURCE_TYPE_ID,
                resource_id,
                assignment["total_hours"],
                start_dt,
                finish_dt,
                1,
                None,
            )

        cur.execute(
            """
            UPDATE jobs
            SET installation_scheduled = True
            WHERE job_id = ?
            """,
            job_id,
        )

        conn.commit()
        conn.close()

        flash(
            "Installation scheduled successfully.",
            "success",
        )

        return redirect(
            url_for(
                "installation_job",
                job_id=job_id,
                stage="install",
                year=year,
                month=month,
            )
        )

    # =============================================================================
    # CAPACITY FOR CALENDAR
    # =============================================================================

    usage_rows = fetch_all(
        """
        SELECT
            resource_id,
            planned_start_date,
            SUM(planned_hours) AS scheduled_hours
        FROM tasks
        WHERE resource_type_id = ?
          AND resource_id IS NOT NULL
          AND planned_start_date IS NOT NULL
        GROUP BY resource_id, planned_start_date
        """,
        [INSTALLATION_RESOURCE_TYPE_ID],
    )

    usage = {}

    for row in usage_rows:

        day_value = (
            row.planned_start_date.date()
            if isinstance(
                row.planned_start_date,
                datetime,
            )
            else row.planned_start_date
        )

        if day_value:
            usage[
                (
                    day_value.strftime("%Y-%m-%d"),
                    row.resource_id,
                )
            ] = float(
                row.scheduled_hours or 0
            )

    available_by_date = {}

    current_day = month_start

    while current_day <= month_end:

        key = current_day.strftime("%Y-%m-%d")

        available_by_date[key] = []

        for resource in resources:
            used = usage.get(
                (
                    key,
                    resource.resource_id,
                ),
                0.0,
            )

            remaining = max(
                0.0,
                INSTALLATION_DAILY_CAPACITY_HOURS - used,
            )

            available_by_date[key].append(
                {
                    "resource_id": resource.resource_id,
                    "resource_name": resource.resource_name,
                    "used_hours": used,
                    "remaining_hours": remaining,
                }
            )

        current_day += timedelta(days=1)

    # =============================================================================
    # EXISTING TASK CHIPS
    # =============================================================================

    scheduled_by_day = {}

    for task in existing_tasks:

        if not task.planned_start_date:
            continue

        task_day = (
            task.planned_start_date.date()
            if isinstance(
                task.planned_start_date,
                datetime,
            )
            else task.planned_start_date
        )

        key = task_day.strftime("%Y-%m-%d")

        scheduled_by_day.setdefault(
            key,
            [],
        ).append(
            {
                "task_name": (
                        task.task_name
                        or f"Task {task.task_id}"
                ),
                "resource_name": (
                        task.resource_name
                        or "Unassigned"
                ),
                "hours": round(
                    float(task.planned_hours or 0),
                    1,
                ),
            }
        )

    # =============================================================================
    # SCREEN
    # =============================================================================

    html = """
    <!doctype html>
    <html>

    <head>
        <title>Schedule Survey / Installation</title>

        <style>

            * {
                box-sizing: border-box;
            }

            body {
                font-family: Arial, sans-serif;
                margin: 0;
                padding: 22px;
                background: #f5f6f8;
            }

            .wrap {
                max-width: 1400px;
                margin: auto;
            }

            .card {
                background: white;
                padding: 22px;
                border-radius: 10px;
                box-shadow: 0 2px 10px rgba(0,0,0,.08);
                margin-bottom: 18px;
            }

            .btn {
                display: inline-block;
                padding: 11px 18px;
                border: 0;
                border-radius: 7px;
                color: white;
                background: #1f4e79;
                text-decoration: none;
                font-weight: bold;
                cursor: pointer;
            }

            .gray {
                background: #6c757d;
            }

            .orange {
                background: #fd7e14;
            }

            .green {
                background: #198754;
            }

            .topbar {
                display: flex;
                justify-content: space-between;
                align-items: center;
            }

            .month-title {
                font-size: 26px;
                font-weight: bold;
                color: #1f4e79;
            }

            .stage-row {
                display: flex;
                gap: 10px;
                flex-wrap: wrap;
                margin-top: 14px;
            }

            .reference-grid {
                display: grid;
                grid-template-columns: repeat(4, 1fr);
                gap: 10px;
                margin-top: 14px;
            }

            .reference {
                background: #f8f9fa;
                padding: 11px;
                border-radius: 7px;
            }

            .cal-grid {
                display: grid;
                grid-template-columns: repeat(7, 1fr);
                gap: 4px;
            }

            .cal-dow {
                text-align: center;
                background: #1f4e79;
                color: white;
                padding: 8px;
                border-radius: 6px;
                font-weight: bold;
            }

            .cal-day {
                min-height: 112px;
                border: 1px solid #ddd;
                border-radius: 6px;
                padding: 6px;
                cursor: pointer;
                background: white;
            }

            .cal-day:hover {
                background: #e8f4e8;
            }

            .cal-day.selected {
                background: #d1e7dd;
                border: 3px solid #198754;
            }

            .cal-day.empty {
                background: #f8f9fa;
                cursor: default;
                border-color: #eee;
            }

            .day-num {
                font-weight: bold;
                margin-bottom: 4px;
            }

            .chip {
                background: #198754;
                color: white;
                font-size: 10px;
                border-radius: 4px;
                padding: 3px 5px;
                margin-bottom: 3px;
                overflow-wrap: anywhere;
            }

            .fields {
                display: grid;
                grid-template-columns: repeat(2, 1fr);
                gap: 12px;
            }

            label {
                display: block;
                font-weight: bold;
                margin-bottom: 5px;
            }

            input,
            select {
                width: 100%;
                padding: 9px;
                border: 1px solid #ccc;
                border-radius: 6px;
            }

            .survey-row {
                display: grid;
                grid-template-columns: 80px 1fr 1fr 1.25fr;
                gap: 10px;
                align-items: center;
                margin-bottom: 9px;
            }


            .resource-checkbox-list {
                display: grid;
                grid-template-columns: repeat(2, minmax(250px, 1fr));
                gap: 8px;
                padding: 12px;
                border: 1px solid #ccc;
                border-radius: 8px;
                background: #fafafa;
            }

            .resource-checkbox-item {
                display: flex;
                align-items: center;
                gap: 8px;
                padding: 8px 10px;
                background: white;
                border: 1px solid #ddd;
                border-radius: 6px;
                cursor: pointer;
                font-weight: normal;
                margin: 0;
            }

            .resource-checkbox-item:hover {
                background: #eef4fb;
            }

            .resource-checkbox-item input[type="checkbox"] {
                width: 18px;
                height: 18px;
                margin: 0;
                cursor: pointer;
                flex: 0 0 auto;
            }

            .resource-grid {
                display: grid;
                grid-template-columns:
                    repeat(
                        auto-fit,
                        minmax(250px, 1fr)
                    );
                gap: 12px;
            }

            .resource-card {
                border: 1px solid #ddd;
                border-radius: 8px;
                padding: 12px;
                background: #fafafa;
            }

            .resource-row {
                padding-bottom: 9px;
                margin-bottom: 9px;
                border-bottom: 1px solid #ddd;
            }

            .flash-success {
                background: #d1e7dd;
                color: #0f5132;
                padding: 10px;
                border-radius: 6px;
                margin: 12px 0;
            }

            .flash-error {
                background: #f8d7da;
                color: #842029;
                padding: 10px;
                border-radius: 6px;
                margin: 12px 0;
            }

            .muted {
                color: #666;
            }

            @media(max-width: 900px) {

                .fields,
                .reference-grid,
                .resource-checkbox-list {
                    grid-template-columns: 1fr;
                }

            }

        </style>
    </head>

    <body>

    <div class="wrap">

        <a
            class="btn gray"
            href="{{ url_for('installation_calendar') }}"
        >
            ← Installation Dashboard
        </a>


        {% with messages = get_flashed_messages(with_categories=true) %}

            {% if messages %}

                {% for category, message in messages %}

                    <div class="flash-{{ category }}">
                        {{ message }}
                    </div>

                {% endfor %}

            {% endif %}

        {% endwith %}


        <div class="card">

            <h1>
                Job {{ header.job_id }}
            </h1>

            <strong>
                {{ header.customer_name or '—' }}
            </strong>


            <div class="reference-grid">

                <div class="reference">

                    <strong>
                        Manufacturing Due
                    </strong>

                    <br>

                    {{
                        header.manufacturing_due_date.strftime('%b %d, %Y')
                        if header.manufacturing_due_date
                        else '—'
                    }}

                </div>


                <div class="reference">

                    <strong>
                        Project Due
                    </strong>

                    <br>

                    {{
                        header.due_date.strftime('%b %d, %Y')
                        if header.due_date
                        else '—'
                    }}

                </div>


                <div class="reference">

                    <strong>
                        Survey
                    </strong>

                    <br>

                    {{
                        'Scheduled'
                        if header.survey_scheduled
                        else 'Not Scheduled'
                    }}

                </div>


                <div class="reference">

                    <strong>
                        Installation
                    </strong>

                    <br>

                    {{
                        'Scheduled'
                        if header.installation_scheduled
                        else 'Not Scheduled'
                    }}

                </div>

            </div>


            <div class="stage-row">

                <a
                    class="btn {% if stage == 'survey' %}orange{% endif %}"
                    href="{{ url_for(
                        'installation_job',
                        job_id=header.job_id,
                        stage='survey',
                        year=year,
                        month=month
                    ) }}"
                >
                    Schedule Survey
                </a>


                <a
                    class="btn {% if stage == 'install' %}green{% endif %}"
                    href="{{ url_for(
                        'installation_job',
                        job_id=header.job_id,
                        stage='install',
                        year=year,
                        month=month
                    ) }}"
                >
                    Schedule Installation
                </a>

            </div>

        </div>


        <div class="card">

            <div class="topbar">

                <a
                    class="btn"
                    href="{{ url_for(
                        'installation_job',
                        job_id=header.job_id,
                        stage=stage,
                        year=prev_year,
                        month=prev_month
                    ) }}"
                >
                    ‹ Prev
                </a>


                <div class="month-title">
                    {{ month_name }}
                </div>


                <a
                    class="btn"
                    href="{{ url_for(
                        'installation_job',
                        job_id=header.job_id,
                        stage=stage,
                        year=next_year,
                        month=next_month
                    ) }}"
                >
                    Next ›
                </a>

            </div>

        </div>


        <form method="post">

            <input
                type="hidden"
                name="stage"
                value="{{ stage }}"
            >

            <input
                type="hidden"
                name="selected_dates"
                id="selected_dates"
            >


            <div class="card">

                <div class="cal-grid">

                    {% for dow in [
                        'Sun',
                        'Mon',
                        'Tue',
                        'Wed',
                        'Thu',
                        'Fri',
                        'Sat'
                    ] %}

                        <div class="cal-dow">
                            {{ dow }}
                        </div>

                    {% endfor %}


                    {% for _ in range(sun_offset) %}

                        <div class="cal-day empty"></div>

                    {% endfor %}


                    {% for d in range(
                        1,
                        days_in_month + 1
                    ) %}

                        {% set day_date =
                            "%04d-%02d-%02d"
                            |format(
                                year,
                                month,
                                d
                            )
                        %}

                        {% set chips =
                            scheduled_by_day.get(
                                day_date,
                                []
                            )
                        %}


                        <div
                            class="cal-day selectable"
                            data-date="{{ day_date }}"
                        >

                            <div class="day-num">
                                {{ d }}
                            </div>


                            {% for chip in chips %}

                                <div class="chip">

                                    {{ chip.task_name }}

                                    <br>

                                    {{ chip.resource_name }}
                                    —
                                    {{ chip.hours }}h

                                </div>

                            {% endfor %}

                        </div>

                    {% endfor %}

                </div>

            </div>


            {% if stage == 'survey' %}


                <div class="card">

                    <h2>
                        Survey Setup
                    </h2>

                    <p class="muted">
                        Click a date to select it.
                        Shift + Click additional dates
                        to schedule surveys across
                        multiple days.
                    </p>


                    <div class="fields">

                        <div>

                            <label>
                                Survey Resources
                            </label>

                            <div class="resource-checkbox-list">

                                {% for resource in resources %}

                                    <label class="resource-checkbox-item">

                                        <input
                                            type="checkbox"
                                            name="survey_resource_ids"
                                            value="{{ resource.resource_id }}"
                                        >

                                        <span>
                                            {{ resource.resource_name }}
                                        </span>

                                    </label>

                                {% endfor %}

                            </div>

                        </div>


                        <div>

                            <label>
                                Number of Surveys
                            </label>

                            <input
                                id="survey_count"
                                type="number"
                                name="survey_count"
                                min="1"
                                step="1"
                                value="1"
                                required
                                oninput="buildSurveyRows()"
                            >

                        </div>

                    </div>

                </div>


                <div class="card">

                    <h2>
                        Survey Types
                    </h2>

                    <div id="surveyRows"></div>

                </div>


                <button
                    class="btn orange"
                    type="submit"
                    style="margin-bottom:40px;"
                >
                    Create Survey Tasks
                </button>


            {% else %}


                <div class="card">

                    <h2>
                        Available Installation Resources
                    </h2>

                    <p class="muted">
                        Click a date to select it.
                        Shift + Click additional dates,
                        then assign crew hours.
                    </p>

                    <div
                        id="resourceAssignments"
                        class="resource-grid"
                    ></div>

                </div>


                <button
                    class="btn green"
                    type="submit"
                    style="margin-bottom:40px;"
                >
                    Schedule Installation
                </button>


            {% endif %}

        </form>

    </div>


    <script>

        const stage =
            {{ stage|tojson }};

        const availability =
            {{ available_by_date|tojson }};

        const qcForms =
            {{ qc_form_options|tojson }};

        const surveyCostCodes =
            {{ survey_cost_code_options|tojson }};

        const selectedDates =
            new Set();


        // ================================================================
        // CALENDAR DATE SELECTION
        //
        // Normal click:
        //      clears previous selections and selects this date.
        //
        // Shift + Click:
        //      adds/removes dates without clearing the others.
        // ================================================================

        document
            .querySelectorAll(".selectable")
            .forEach(function(day) {

                day.addEventListener(
                    "click",
                    function(event) {

                        const value =
                            this.dataset.date;


                        if (event.shiftKey) {

                            if (
                                selectedDates.has(value)
                            ) {

                                selectedDates.delete(
                                    value
                                );

                            } else {

                                selectedDates.add(
                                    value
                                );

                            }

                        } else {

                            selectedDates.clear();

                            selectedDates.add(
                                value
                            );

                        }


                        refreshSelection();

                    }
                );

            });


        function refreshSelection() {

            document
                .querySelectorAll(
                    ".selectable"
                )
                .forEach(function(day) {

                    day.classList.toggle(
                        "selected",
                        selectedDates.has(
                            day.dataset.date
                        )
                    );

                });


            const dates =
                Array.from(
                    selectedDates
                ).sort();


            document.getElementById(
                "selected_dates"
            ).value =
                dates.join(",");


            if (
                stage === "install"
            ) {

                buildResourceAssignments();

            }

        }


        // ================================================================
        // SURVEY QC FORM ROWS
        // ================================================================

        function buildSurveyRows() {

            const holder =
                document.getElementById(
                    "surveyRows"
                );

            const input =
                document.getElementById(
                    "survey_count"
                );


            if (
                !holder ||
                !input
            ) {
                return;
            }


            const count =
                parseInt(
                    input.value || "0"
                );

            let costCodeOptions =
                `<option value="">
                    -- Select Cost Code --
                </option>`;

            surveyCostCodes.forEach(
                function(code) {

                    let label =
                        code.coins_cost_code ||
                        `Cost Code ${code.cost_code_id}`;

                    if (code.cost_code_description) {
                        label +=
                            ` — ${code.cost_code_description}`;
                    }

                    if (code.line_item_description) {
                        label +=
                            ` — ${code.line_item_description}`;
                    }

                    costCodeOptions += `
                        <option value="${code.cost_code_id}">
                            ${label}
                        </option>
                    `;
                }
            );

            holder.innerHTML = "";


            for (
                let i = 1;
                i <= count;
                i++
            ) {

                let options =
                    `<option value="">
                        -- Select QC Form --
                    </option>`;


                qcForms.forEach(
                    function(form) {

                        options += `
                            <option
                                value="${form.qc_form_id}"
                            >
                                ${form.form_name}
                            </option>
                        `;

                    }
                );


               holder.innerHTML += `
    <div class="survey-row">

        <strong>
            #${i}
        </strong>

        <select
            name="survey_form_${i}"
            required
        >
            ${options}
        </select>

        <input
            type="text"
            name="survey_suffix_${i}"
            placeholder="Task text, e.g. Main Entrance"
            required
        >

        <select
            name="survey_cost_code_${i}"
            required
        >
            ${costCodeOptions}
        </select>

    </div>
`;

            }

        }


        // ================================================================
        // INSTALLATION RESOURCE ASSIGNMENTS
        // ================================================================

        function buildResourceAssignments() {

            const holder =
                document.getElementById(
                    "resourceAssignments"
                );


            if (!holder) {
                return;
            }


            holder.innerHTML = "";


            Array
                .from(
                    selectedDates
                )
                .sort()
                .forEach(
                    function(dateValue) {

                        const card =
                            document.createElement(
                                "div"
                            );

                        card.className =
                            "resource-card";


                        let html = `
                            <strong>
                                ${dateValue}
                            </strong>

                            <br><br>
                        `;


                        const resourcesForDate =
                            availability[
                                dateValue
                            ] || [];


                        resourcesForDate.forEach(
                            function(resource) {

                                html += `
                                    <div
                                        class="resource-row"
                                    >

                                        <strong>
                                            ${resource.resource_name}
                                        </strong>

                                        <br>

                                        <small>
                                            ${resource.remaining_hours.toFixed(1)}
                                            hrs available
                                        </small>

                                        <input
                                            type="number"
                                            step="0.5"
                                            min="0"
                                            max="${resource.remaining_hours}"
                                            name="hours_${dateValue}_${resource.resource_id}"
                                            placeholder="Hours to assign"
                                        >

                                    </div>
                                `;

                            }
                        );


                        card.innerHTML =
                            html;

                        holder.appendChild(
                            card
                        );

                    }
                );

        }


        buildSurveyRows();

        refreshSelection();

    </script>

    </body>
    </html>
    """

    return render_template_string(
        html,
        header=header,
        stage=stage,
        month_name=month_name,
        year=year,
        month=month,
        days_in_month=days_in_month,
        sun_offset=sun_offset,
        prev_year=prev_year,
        prev_month=prev_month,
        next_year=next_year,
        next_month=next_month,
        resources=resources,
        qc_form_options=qc_form_options,
        survey_cost_code_options=survey_cost_code_options,
        available_by_date=available_by_date,
        scheduled_by_day=scheduled_by_day,
    )
########################################################################################################################
########################################################################################################################
########################################################################################################################





########################################################################################################################
########################################################################################################################
########################################################################################################################
"""RESOURCE MANAGEMENT"""
########################################################################################################################
########################################################################################################################
########################################################################################################################
@app.route("/admin/resource-management", methods=["GET", "POST"])
@login_required
@office_required
def resource_management():
    employee = get_logged_in_employee()
    if not _can_manage_resources(employee):
        flash("Access restricted.", "error")
        return redirect(url_for("home"))

    action = request.form.get("action", "").strip() if request.method == "POST" else ""

    # -----------------------------------------------------------------------
    # Resource actions
    # -----------------------------------------------------------------------
    if request.method == "POST" and action == "add_resource":
        resource_name = request.form.get("resource_name_new", "").strip()
        resource_type_id = _parse_int(request.form.get("resource_type_new"), 0)
        default_hours = _parse_float(request.form.get("default_hours_new"), 0)

        if not resource_name:
            flash("Resource name is required.", "error")
        elif resource_type_id <= 0:
            flash("Please select a resource type.", "error")
        elif default_hours <= 0:
            flash("Default hours per day must be greater than zero.", "error")
        else:
            try:
                execute_sql(
                    """
                    INSERT INTO resources
                        (resource_name, resource_types, default_hours_per_day, resource_members)
                    VALUES (?, ?, ?, ?)
                    """,
                    [resource_name, resource_type_id, default_hours, ""],
                )
                flash(f"Resource '{resource_name}' added.", "success")
            except Exception as ex:
                flash(f"Could not add resource: {ex}", "error")

        return redirect(url_for("resource_management") + "#resources")

    if request.method == "POST" and action.startswith("save_resource:"):
        resource_id = _parse_int(action.split(":", 1)[1], 0)
        resource_name = request.form.get(f"resource_name_{resource_id}", "").strip()
        resource_type_id = _parse_int(request.form.get(f"resource_type_{resource_id}"), 0)
        default_hours = _parse_float(request.form.get(f"default_hours_{resource_id}"), 0)

        if resource_id <= 0:
            flash("Invalid resource.", "error")
        elif not resource_name:
            flash("Resource name is required.", "error")
        elif resource_type_id <= 0:
            flash("Please select a resource type.", "error")
        elif default_hours <= 0:
            flash("Default hours per day must be greater than zero.", "error")
        else:
            try:
                execute_sql(
                    """
                    UPDATE resources
                    SET resource_name = ?,
                        resource_types = ?,
                        default_hours_per_day = ?
                    WHERE resource_id = ?
                    """,
                    [resource_name, resource_type_id, default_hours, resource_id],
                )
                flash(f"Resource {resource_id} saved.", "success")
            except Exception as ex:
                flash(f"Could not save resource: {ex}", "error")

        return redirect(url_for("resource_management") + "#resources")

    if request.method == "POST" and action == "sync_members":
        try:
            _sync_resource_members_from_employees()
            flash("Resource membership synchronized from employee assignments.", "success")
        except Exception as ex:
            flash(f"Could not synchronize resource members: {ex}", "error")
        return redirect(url_for("resource_management") + "#resources")

    # -----------------------------------------------------------------------
    # Employee actions
    # -----------------------------------------------------------------------
    if request.method == "POST" and action == "add_employee":
        first_name = request.form.get("employee_first_name_new", "").strip()
        last_name = request.form.get("employee_last_name_new", "").strip()
        resource_id = _parse_int(request.form.get("employee_resource_new"), 0)
        email = request.form.get("email_address_new", "").strip()
        password = request.form.get("password_new", "").strip()

        office = -1 if request.form.get("office_new") == "1" else 0
        installer = -1 if request.form.get("installer_new") == "1" else 0
        admin = -1 if request.form.get("admin_new") == "1" else 0

        if not first_name or not last_name:
            flash("First and last name are required.", "error")
        elif not email:
            flash("Email address is required.", "error")
        elif not password:
            flash("A password is required for a new employee.", "error")
        else:
            try:
                execute_sql(
                    """
                    INSERT INTO employees
                        (employee_first_name, employee_last_name,
                         resource_id, email_address, [password],
                         [office], [installer], [admin])
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    [
                        first_name, last_name, resource_id, email, password,
                        office, installer, admin,
                    ],
                )
                _sync_resource_members_from_employees()
                flash(f"Employee {first_name} {last_name} added.", "success")
            except Exception as ex:
                flash(f"Could not add employee: {ex}", "error")

        return redirect(url_for("resource_management") + "#employees")

    if request.method == "POST" and action.startswith("save_employee:"):
        employee_id = _parse_int(action.split(":", 1)[1], 0)
        first_name = request.form.get(f"employee_first_name_{employee_id}", "").strip()
        last_name = request.form.get(f"employee_last_name_{employee_id}", "").strip()
        resource_id = _parse_int(request.form.get(f"employee_resource_{employee_id}"), 0)
        email = request.form.get(f"email_address_{employee_id}", "").strip()
        new_password = request.form.get(f"password_{employee_id}", "").strip()

        office = -1 if request.form.get(f"office_{employee_id}") == "1" else 0
        installer = -1 if request.form.get(f"installer_{employee_id}") == "1" else 0
        admin = -1 if request.form.get(f"admin_{employee_id}") == "1" else 0

        if employee_id <= 0:
            flash("Invalid employee.", "error")
        elif not first_name or not last_name:
            flash("First and last name are required.", "error")
        elif not email:
            flash("Email address is required.", "error")
        else:
            try:
                if new_password:
                    execute_sql(
                        """
                        UPDATE employees
                        SET employee_first_name = ?,
                            employee_last_name = ?,
                            resource_id = ?,
                            email_address = ?,
                            [password] = ?,
                            [office] = ?,
                            [installer] = ?,
                            [admin] = ?
                        WHERE employee_id = ?
                        """,
                        [
                            first_name, last_name, resource_id, email,
                            new_password, office, installer, admin,
                            employee_id,
                        ],
                    )
                else:
                    execute_sql(
                        """
                        UPDATE employees
                        SET employee_first_name = ?,
                            employee_last_name = ?,
                            resource_id = ?,
                            email_address = ?,
                            [office] = ?,
                            [installer] = ?,
                            [admin] = ?
                        WHERE employee_id = ?
                        """,
                        [
                            first_name, last_name, resource_id, email,
                            office, installer, admin, employee_id,
                        ],
                    )

                _sync_resource_members_from_employees()
                flash(f"Employee {first_name} {last_name} saved.", "success")
            except Exception as ex:
                flash(f"Could not save employee: {ex}", "error")

        return redirect(url_for("resource_management") + "#employees")

    # -----------------------------------------------------------------------
    # Page data - loaded separately to avoid Access JOIN type-mismatch issues.
    # -----------------------------------------------------------------------
    resource_type_rows = fetch_all(
        """
        SELECT resource_type_id, resource_description
        FROM resource_types
        ORDER BY resource_description
        """,
        [],
    )

    resource_rows = fetch_all(
        """
        SELECT resource_id, resource_name, resource_types,
               default_hours_per_day, resource_members
        FROM resources
        ORDER BY resource_name
        """,
        [],
    )

    employee_rows = fetch_all(
        """
        SELECT employee_id,
               employee_first_name, employee_last_name,
               resource_id, email_address,
               [office], [installer], [admin]
        FROM employees
        ORDER BY employee_first_name, employee_last_name
        """,
        [],
    )

    from types import SimpleNamespace

    resource_types = [
        SimpleNamespace(
            resource_type_id=rt.resource_type_id,
            resource_type_key=_access_id_key(rt.resource_type_id),
            resource_description=rt.resource_description or "",
        )
        for rt in resource_type_rows
    ]

    # Derive membership directly from employees so the screen always reflects
    # the employee.resource_id source of truth, even before a manual sync.
    employees_by_resource = {}
    for emp in employee_rows:
        key = _access_id_key(emp.resource_id)
        if not key:
            continue
        name = f"{emp.employee_first_name or ''} {emp.employee_last_name or ''}".strip()
        employees_by_resource.setdefault(key, []).append(
            (emp.employee_id, name)
        )

    resources = []
    for r in resource_rows:
        key = _access_id_key(r.resource_id)
        members = employees_by_resource.get(key, [])
        resources.append(
            SimpleNamespace(
                resource_id=r.resource_id,
                resource_key=key,
                resource_name=r.resource_name or "",
                resource_types=r.resource_types,
                resource_type_key=_access_id_key(r.resource_types),
                default_hours_per_day=r.default_hours_per_day if r.default_hours_per_day is not None else "",
                stored_resource_members=r.resource_members or "",
                members=members,
                member_count=len(members),
            )
        )

    employees = []
    for emp in employee_rows:
        employees.append(
            SimpleNamespace(
                employee_id=emp.employee_id,
                employee_first_name=emp.employee_first_name or "",
                employee_last_name=emp.employee_last_name or "",
                resource_id=emp.resource_id,
                resource_key=_access_id_key(emp.resource_id),
                email_address=emp.email_address or "",
                office_bool=is_yes_no_true(emp.office),
                installer_bool=is_yes_no_true(emp.installer),
                admin_bool=is_yes_no_true(emp.admin),
            )
        )

    html = """
    <!doctype html>
    <html>
    <head>
        <title>Resource Management</title>
        <style>
            * { box-sizing: border-box; }
            body {
                font-family: Arial, sans-serif;
                margin: 0;
                background: #f5f6f8;
                padding: 18px 22px 40px;
                color: #222;
            }
            .wrap { max-width: 1700px; margin: 0 auto; }
            .card {
                background: white;
                padding: 22px;
                border-radius: 10px;
                box-shadow: 0 2px 10px rgba(0,0,0,.08);
                margin-bottom: 18px;
            }
            .topbar {
                display: flex;
                align-items: center;
                justify-content: space-between;
                gap: 16px;
            }
            h1 { color: #1f4e79; margin: 0; font-size: 26px; }
            h2 { color: #1f4e79; margin: 0 0 14px; font-size: 18px; }
            .subtitle { color: #666; font-size: 13px; margin-top: 5px; }
            .btn {
                display: inline-block;
                padding: 9px 15px;
                background: #1f4e79;
                color: white;
                text-decoration: none;
                border: none;
                border-radius: 7px;
                font-weight: bold;
                cursor: pointer;
                font-size: 13px;
                white-space: nowrap;
            }
            .btn:hover { opacity: .92; }
            .btn-secondary { background: #6c757d; }
            .btn-light {
                background: #e9eef3;
                color: #1f4e79;
                border: 1px solid #cbd5df;
            }
            .flash-success {
                background: #d1e7dd;
                color: #0f5132;
                padding: 11px 14px;
                border-radius: 7px;
                margin-bottom: 12px;
            }
            .flash-error {
                background: #f8d7da;
                color: #842029;
                padding: 11px 14px;
                border-radius: 7px;
                margin-bottom: 12px;
            }
            .tabs {
                display: flex;
                gap: 8px;
                margin-bottom: 18px;
            }
            .tab-btn {
                padding: 11px 20px;
                border: 1px solid #cbd5df;
                background: white;
                border-radius: 8px;
                color: #1f4e79;
                font-weight: bold;
                cursor: pointer;
            }
            .tab-btn.active {
                background: #1f4e79;
                color: white;
                border-color: #1f4e79;
            }
            .tab-panel { display: none; }
            .tab-panel.active { display: block; }
            .grid-form {
                display: grid;
                grid-template-columns: repeat(5, minmax(140px, 1fr));
                gap: 12px;
                align-items: end;
            }
            .employee-add-grid {
                grid-template-columns: 1fr 1fr 1.4fr 1.6fr 1fr;
            }
            label {
                display: block;
                font-size: 11px;
                font-weight: bold;
                color: #555;
                margin-bottom: 4px;
            }
            input[type=text],
            input[type=email],
            input[type=password],
            input[type=number],
            select {
                width: 100%;
                padding: 7px 8px;
                border: 1px solid #cfd6dd;
                border-radius: 6px;
                font-size: 12px;
                background: white;
            }
            input[type=checkbox] {
                width: 16px;
                height: 16px;
                cursor: pointer;
            }
            .checks {
                display: flex;
                gap: 16px;
                flex-wrap: wrap;
                align-items: center;
                margin-top: 12px;
            }
            .check-label {
                display: flex;
                align-items: center;
                gap: 5px;
                font-size: 12px;
                font-weight: normal;
                color: #333;
                margin: 0;
            }
            .table-tools {
                display: flex;
                justify-content: space-between;
                align-items: center;
                gap: 12px;
                margin-bottom: 10px;
                flex-wrap: wrap;
            }
            .search {
                width: 320px !important;
                max-width: 100%;
            }
            .table-wrap {
                overflow-x: auto;
                border: 1px solid #e2e6ea;
                border-radius: 8px;
            }
            table {
                width: 100%;
                border-collapse: collapse;
                min-width: 1000px;
            }
            .employee-table { min-width: 1250px; }
            th, td {
                padding: 8px 8px;
                border-bottom: 1px solid #e8ebee;
                text-align: left;
                font-size: 12px;
                vertical-align: middle;
            }
            th {
                background: #eef3f7;
                position: sticky;
                top: 0;
                z-index: 1;
                white-space: nowrap;
            }
            tbody tr:nth-child(even) { background: #fafbfc; }
            tbody tr:hover { background: #f2f7fb; }
            td.center, th.center { text-align: center; }
            .id-cell { color: #777; font-weight: bold; width: 55px; }
            .member-list {
                color: #555;
                font-size: 11px;
                line-height: 1.45;
                min-width: 180px;
            }
            .member-count {
                display: inline-block;
                background: #e7f0f8;
                color: #1f4e79;
                padding: 2px 7px;
                border-radius: 999px;
                font-weight: bold;
                margin-right: 6px;
            }
            .note {
                font-size: 12px;
                color: #666;
                line-height: 1.4;
                margin-bottom: 12px;
            }
            .flags {
                display: grid;
                grid-template-columns: repeat(3, auto);
                gap: 10px;
                justify-content: start;
            }
            @media (max-width: 1000px) {
                .grid-form, .employee-add-grid { grid-template-columns: 1fr 1fr; }
            }
        </style>
    </head>
    <body>
    <div class="wrap">

        <div class="card topbar">
            <div>
                <h1>Resource Management</h1>
                <div class="subtitle">Manage resources, employee information, resource assignments, and application roles.</div>
            </div>
            <a class="btn btn-secondary" href="{{ url_for('home') }}">← Back</a>
        </div>

        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}
                {% for category, message in messages %}
                    <div class="flash-{{ category }}">{{ message }}</div>
                {% endfor %}
            {% endif %}
        {% endwith %}

        <div class="tabs">
            <button class="tab-btn" id="tab-resources" type="button" onclick="showTab('resources')">Resources</button>
            <button class="tab-btn" id="tab-employees" type="button" onclick="showTab('employees')">Employees</button>
        </div>

        <!-- ============================================================= -->
        <!-- RESOURCES -->
        <!-- ============================================================= -->
        <div class="tab-panel" id="panel-resources">

            <div class="card">
                <h2>Add Resource</h2>
                <form method="post">
                    <div class="grid-form">
                        <div>
                            <label>Resource Name</label>
                            <input type="text" name="resource_name_new" required>
                        </div>
                        <div>
                            <label>Resource Type</label>
                            <select name="resource_type_new" required>
                                <option value="">-- Select --</option>
                                {% for rt in resource_types %}
                                <option value="{{ rt.resource_type_id }}">{{ rt.resource_description }}</option>
                                {% endfor %}
                            </select>
                        </div>
                        <div>
                            <label>Default Hours / Day</label>
                            <input type="number" name="default_hours_new" min="0.1" step="0.1" value="7" required>
                        </div>
                        <div>
                            <button class="btn" type="submit" name="action" value="add_resource">Add Resource</button>
                        </div>
                    </div>
                </form>
            </div>

            <div class="card">
                <div class="table-tools">
                    <div>
                        <h2 style="margin-bottom:4px;">Existing Resources</h2>
                        <div class="note" style="margin:0;">
                            Members are derived from each employee's assigned resource. The stored
                            <code>resource_members</code> field is synchronized automatically after employee changes.
                        </div>
                    </div>
                    <div style="display:flex;gap:8px;align-items:center;">
                        <input class="search" id="resource-search" type="text"
                               placeholder="Search resources..." oninput="filterRows('resource-table-body', this.value)">
                        <form method="post" style="margin:0;">
                            <button class="btn btn-light" type="submit" name="action" value="sync_members">Sync Members</button>
                        </form>
                    </div>
                </div>

                <form method="post">
                <div class="table-wrap">
                    <table>
                        <thead>
                            <tr>
                                <th>ID</th>
                                <th>Resource Name</th>
                                <th>Resource Type</th>
                                <th>Hours / Day</th>
                                <th>Assigned Employees</th>
                                <th></th>
                            </tr>
                        </thead>
                        <tbody id="resource-table-body">
                        {% for res in resources %}
                            <tr>
                                <td class="id-cell">{{ res.resource_id }}</td>
                                <td>
                                    <input type="text" name="resource_name_{{ res.resource_id }}"
                                           value="{{ res.resource_name }}" required>
                                </td>
                                <td>
                                    <select name="resource_type_{{ res.resource_id }}">
                                        {% for rt in resource_types %}
                                        <option value="{{ rt.resource_type_id }}"
                                            {% if res.resource_type_key == rt.resource_type_key %}selected{% endif %}>
                                            {{ rt.resource_description }}
                                        </option>
                                        {% endfor %}
                                    </select>
                                </td>
                                <td style="width:120px;">
                                    <input type="number" name="default_hours_{{ res.resource_id }}"
                                           min="0.1" step="0.1" value="{{ res.default_hours_per_day }}">
                                </td>
                                <td class="member-list">
                                    <span class="member-count">{{ res.member_count }}</span>
                                    {% if res.members %}
                                        {% for employee_id, name in res.members %}
                                            #{{ employee_id }}{% if name %} ({{ name }}){% endif %}{% if not loop.last %}, {% endif %}
                                        {% endfor %}
                                    {% else %}
                                        No employees assigned
                                    {% endif %}
                                </td>
                                <td style="width:80px;">
                                    <button class="btn" type="submit" name="action"
                                            value="save_resource:{{ res.resource_id }}">Save</button>
                                </td>
                            </tr>
                        {% endfor %}
                        </tbody>
                    </table>
                </div>
                </form>
            </div>
        </div>

        <!-- ============================================================= -->
        <!-- EMPLOYEES -->
        <!-- ============================================================= -->
        <div class="tab-panel" id="panel-employees">

            <div class="card">
                <h2>Add Employee</h2>
                <form method="post">
                    <div class="grid-form employee-add-grid">
                        <div>
                            <label>First Name</label>
                            <input type="text" name="employee_first_name_new" required>
                        </div>
                        <div>
                            <label>Last Name</label>
                            <input type="text" name="employee_last_name_new" required>
                        </div>
                        <div>
                            <label>Resource</label>
                            <select name="employee_resource_new">
                                <option value="0">-- None / Office --</option>
                                {% for res in resources %}
                                <option value="{{ res.resource_id }}">{{ res.resource_name }}</option>
                                {% endfor %}
                            </select>
                        </div>
                        <div>
                            <label>Email Address</label>
                            <input type="email" name="email_address_new" required>
                        </div>
                        <div>
                            <label>Initial Password</label>
                            <input type="password" name="password_new" required>
                        </div>
                    </div>

                    <div class="checks">
                        <label class="check-label"><input type="checkbox" name="office_new" value="1"> Office</label>
                        <label class="check-label"><input type="checkbox" name="installer_new" value="1"> Installer</label>
                        <label class="check-label"><input type="checkbox" name="admin_new" value="1"> Admin</label>
                        <button class="btn" type="submit" name="action" value="add_employee">Add Employee</button>
                    </div>
                </form>
            </div>

            <div class="card">
                <div class="table-tools">
                    <div>
                        <h2 style="margin-bottom:4px;">Employees</h2>
                        <div class="note" style="margin:0;">
                            Change an employee's Resource here to move them between crews/departments.
                            Leave New Password blank when saving to keep the current password.
                        </div>
                    </div>
                    <input class="search" id="employee-search" type="text"
                           placeholder="Search employees..." oninput="filterRows('employee-table-body', this.value)">
                </div>

                <form method="post">
                <div class="table-wrap">
                    <table class="employee-table">
                        <thead>
                            <tr>
                                <th>ID</th>
                                <th>First Name</th>
                                <th>Last Name</th>
                                <th>Resource</th>
                                <th>Email</th>
                                <th>New Password</th>
                                <th class="center">Office</th>
                                <th class="center">Installer</th>
                                <th class="center">Admin</th>
                                <th></th>
                            </tr>
                        </thead>
                        <tbody id="employee-table-body">
                        {% for emp in employees %}
                            <tr>
                                <td class="id-cell">{{ emp.employee_id }}</td>
                                <td>
                                    <input type="text" name="employee_first_name_{{ emp.employee_id }}"
                                           value="{{ emp.employee_first_name }}" required>
                                </td>
                                <td>
                                    <input type="text" name="employee_last_name_{{ emp.employee_id }}"
                                           value="{{ emp.employee_last_name }}" required>
                                </td>
                                <td style="min-width:210px;">
                                    <select name="employee_resource_{{ emp.employee_id }}">
                                        <option value="0"
                                            {% if not emp.resource_key or emp.resource_key == '0' %}selected{% endif %}>
                                            -- None / Office --
                                        </option>
                                        {% for res in resources %}
                                        <option value="{{ res.resource_id }}"
                                            {% if emp.resource_key == res.resource_key %}selected{% endif %}>
                                            {{ res.resource_name }}
                                        </option>
                                        {% endfor %}
                                    </select>
                                </td>
                                <td style="min-width:210px;">
                                    <input type="email" name="email_address_{{ emp.employee_id }}"
                                           value="{{ emp.email_address }}" required>
                                </td>
                                <td style="width:145px;">
                                    <input type="password" name="password_{{ emp.employee_id }}"
                                           placeholder="Keep current">
                                </td>
                                <td class="center">
                                    <input type="checkbox" name="office_{{ emp.employee_id }}" value="1"
                                           {% if emp.office_bool %}checked{% endif %}>
                                </td>
                                <td class="center">
                                    <input type="checkbox" name="installer_{{ emp.employee_id }}" value="1"
                                           {% if emp.installer_bool %}checked{% endif %}>
                                </td>
                                <td class="center">
                                    <input type="checkbox" name="admin_{{ emp.employee_id }}" value="1"
                                           {% if emp.admin_bool %}checked{% endif %}>
                                </td>
                                <td style="width:80px;">
                                    <button class="btn" type="submit" name="action"
                                            value="save_employee:{{ emp.employee_id }}">Save</button>
                                </td>
                            </tr>
                        {% endfor %}
                        </tbody>
                    </table>
                </div>
                </form>
            </div>
        </div>
    </div>

    <script>
        function showTab(name) {
            document.querySelectorAll('.tab-panel').forEach(function(el) {
                el.classList.remove('active');
            });
            document.querySelectorAll('.tab-btn').forEach(function(el) {
                el.classList.remove('active');
            });

            var panel = document.getElementById('panel-' + name);
            var button = document.getElementById('tab-' + name);
            if (panel) panel.classList.add('active');
            if (button) button.classList.add('active');

            if (history.replaceState) {
                history.replaceState(null, '', '#' + name);
            }
        }

        function filterRows(tbodyId, query) {
            var q = (query || '').toLowerCase().trim();
            var tbody = document.getElementById(tbodyId);
            if (!tbody) return;
            Array.from(tbody.querySelectorAll('tr')).forEach(function(row) {
                row.style.display = row.textContent.toLowerCase().includes(q) ? '' : 'none';
            });
        }

        document.addEventListener('DOMContentLoaded', function() {
            var requested = (window.location.hash || '').replace('#', '');
            if (requested !== 'employees' && requested !== 'resources') {
                requested = 'resources';
            }
            showTab(requested);
        });
    </script>
    </body>
    </html>
    """

    return render_template_string(
        html,
        resource_types=resource_types,
        resources=resources,
        employees=employees,
    )
########################################################################################################################
########################################################################################################################
########################################################################################################################





########################################################################################################################
########################################################################################################################
########################################################################################################################
"""MANUFACTURING FLOOR DISPLAY"""
########################################################################################################################
########################################################################################################################
########################################################################################################################
@app.route("/manufacturing-floor-display")
@login_required
def manufacturing_floor_display():
    """
    Shop-floor task board.

    Three columns:
      1. Tasks Not Started
      2. Tasks In Process
      3. Tasks Completed Today

    Filters:
      Resource Type -> Resource Name
    """
    from collections import defaultdict

    now = datetime.now()
    today = now.date()

    resource_type_filter = request.args.get("resource_type", "").strip()
    resource_filter = request.args.get("resource", "").strip()

    # Show ALL resource types and resources, including Installation.
    resource_types = list(get_resource_types())
    resources = list(get_resources())

    resource_type_name_map = {
        int(rt.resource_type_id): (rt.resource_description or "")
        for rt in resource_types
        if rt.resource_type_id is not None
    }
    resource_name_map = {
        int(r.resource_id): (r.resource_name or "")
        for r in resources
        if r.resource_id is not None
    }
    resource_type_by_resource = {
        int(r.resource_id): int(r.resource_types or 0)
        for r in resources
        if r.resource_id is not None
    }
    status_name_map = get_status_name_map()

    selected_resource_type_id = None
    selected_resource_id = None

    try:
        if resource_type_filter:
            selected_resource_type_id = int(resource_type_filter)
    except (TypeError, ValueError):
        resource_type_filter = ""

    try:
        if resource_filter:
            selected_resource_id = int(resource_filter)
    except (TypeError, ValueError):
        resource_filter = ""

    # Resource Name is subordinate to Resource Type.
    # If the selected resource does not belong to the selected type,
    # clear the resource selection.
    if selected_resource_id is not None:
        selected_type_for_resource = resource_type_by_resource.get(selected_resource_id)

        if (
            selected_resource_type_id is None
            or selected_type_for_resource != selected_resource_type_id
        ):
            selected_resource_id = None
            resource_filter = ""

    resource_options = []

    if selected_resource_type_id is not None:
        resource_options = [
            r
            for r in resources
            if int(r.resource_types or 0) == selected_resource_type_id
        ]
        resource_options.sort(
            key=lambda r: (r.resource_name or "").lower()
        )

    tasks = fetch_all(
        """
        SELECT
            t.task_id,
            t.task_name,
            t.resource_type_id,
            t.resource_id,
            t.status_id,
            t.sequence_number,
            t.planned_hours,
            t.actual_start_date,
            t.actual_finish_date,
            li.job_number_id,
            li.work_order_id
        FROM
            (tasks AS t
            LEFT JOIN cost_codes AS cc
                ON t.cost_code_id = cc.cost_code_id)
            LEFT JOIN line_items AS li
                ON cc.line_item_id = li.line_item_id
        WHERE
            t.resource_type_id IS NOT NULL
            AND t.task_id NOT IN (?, ?)
        ORDER BY
            t.sequence_number,
            t.task_id
        """,
        [TECH_OPS_TASK_ID, OVERHEAD_TASK_ID],
    )

    jobs_raw = fetch_all(
        """
        SELECT
            j.job_id,
            j.due_date,
            c.customer_name
        FROM jobs AS j
        LEFT JOIN customers AS c
            ON j.customer_id = c.customer_id
        """
    )

    work_orders_raw = fetch_all(
        """
        SELECT
            w.wo_id,
            w.due_date,
            c.customer_name
        FROM work_orders AS w
        LEFT JOIN customers AS c
            ON w.customer_id = c.customer_id
        """
    )

    jobs_map = {
        int(row.job_id): {
            "customer": row.customer_name or "",
            "due_date": row.due_date,
        }
        for row in jobs_raw
    }

    work_orders_map = {
        int(row.wo_id): {
            "customer": row.customer_name or "",
            "due_date": row.due_date,
        }
        for row in work_orders_raw
    }

    punches = fetch_all(
        """
        SELECT
            tp.task_id,
            tp.employee_id,
            tp.punch_in,
            tp.punch_out,
            tp.task_time,
            e.employee_first_name,
            e.employee_last_name
        FROM time_punches AS tp
        LEFT JOIN employees AS e
            ON tp.employee_id = e.employee_id
        WHERE tp.task_id IS NOT NULL
        """
    )

    task_minutes = defaultdict(float)
    active_employees = defaultdict(list)

    for punch in punches:
        if punch.task_id is None or not punch.punch_in:
            continue

        if punch.punch_out is None:
            minutes = max(
                (now - punch.punch_in).total_seconds() / 60.0,
                0,
            )

            employee_name = (
                f"{punch.employee_first_name or ''} "
                f"{punch.employee_last_name or ''}"
            ).strip()

            if (
                employee_name
                and employee_name not in active_employees[punch.task_id]
            ):
                active_employees[punch.task_id].append(employee_name)

        elif punch.task_time is not None:
            minutes = float(punch.task_time or 0)

        else:
            minutes = max(
                (punch.punch_out - punch.punch_in).total_seconds() / 60.0,
                0,
            )

        task_minutes[punch.task_id] += minutes

    def normalize_status(status_id):
        try:
            return int(status_id)
        except (TypeError, ValueError):
            return 0

    def is_complete_status(status_value):
        return status_value == 3

    def is_in_process_status(status_value):
        return status_value == 2

    def same_day(value, day):
        if value is None:
            return False

        try:
            return value.date() == day
        except AttributeError:
            return value == day

    def fmt_mins(minutes):
        minutes = int(round(float(minutes or 0)))
        hours, mins = divmod(minutes, 60)

        if hours and mins:
            return f"{hours}h {mins}m"

        if hours:
            return f"{hours}h"

        if mins:
            return f"{mins}m"

        return "—"

    def fmt_hours(hours):
        if hours is None:
            return "—"

        try:
            return f"{float(hours):.1f}h"
        except (TypeError, ValueError):
            return "—"

    columns = {
        "not_started": [],
        "in_process": [],
        "completed_today": [],
    }

    for task in tasks:
        resource_type_id = int(task.resource_type_id or 0)

        resource_id = (
            int(task.resource_id or 0)
            if task.resource_id
            else None
        )

        if (
            selected_resource_type_id is not None
            and resource_type_id != selected_resource_type_id
        ):
            continue

        if (
            selected_resource_id is not None
            and resource_id != selected_resource_id
        ):
            continue

        job_number_id = (
            int(task.job_number_id or 0)
            if task.job_number_id
            else None
        )

        work_order_id = (
            int(task.work_order_id or 0)
            if task.work_order_id
            else None
        )

        if job_number_id:
            entity_ref = f"Job {job_number_id}"
            entity_info = jobs_map.get(job_number_id, {})

        elif work_order_id:
            entity_ref = f"WO {work_order_id}"
            entity_info = work_orders_map.get(work_order_id, {})

        else:
            entity_ref = "No Job / WO"
            entity_info = {}

        status_value = normalize_status(task.status_id)
        active_names = list(
            active_employees.get(task.task_id, [])
        )

        completed_today = (
                is_complete_status(status_value)
                and same_day(task.actual_finish_date, today)
        )

        if completed_today:
            bucket = "completed_today"

        elif status_value == 2 or active_names:
            bucket = "in_process"

        elif status_value == 3:
            # Completed before today should not remain on the board.
            continue

        elif status_value == 1:
            bucket = "not_started"

        else:
            continue

        columns[bucket].append(
            {
                "task_id": int(task.task_id),
                "task_name": task.task_name or f"Task {task.task_id}",
                "entity_ref": entity_ref,
                "customer": entity_info.get("customer", ""),
                "due_date": entity_info.get("due_date"),
                "resource_type": resource_type_name_map.get(
                    resource_type_id,
                    "—",
                ),
                "resource_name": (
                    resource_name_map.get(resource_id, "Unassigned")
                    if resource_id
                    else "Unassigned"
                ),
                "planned_hours": task.planned_hours,
                "logged_minutes": task_minutes.get(task.task_id, 0),
                "active_names": active_names,
                "actual_start_date": task.actual_start_date,
                "actual_finish_date": task.actual_finish_date,
                "sequence_number": int(
                    task.sequence_number or 9999
                ),
            }
        )

    far_future = datetime(2099, 12, 31)

    columns["not_started"].sort(
        key=lambda row: (
            row["due_date"] or far_future,
            row["sequence_number"],
            row["entity_ref"],
            row["task_name"],
        )
    )

    columns["in_process"].sort(
        key=lambda row: (
            row["due_date"] or far_future,
            row["actual_start_date"] or far_future,
            row["entity_ref"],
            row["task_name"],
        )
    )

    columns["completed_today"].sort(
        key=lambda row: (
            row["actual_finish_date"] or datetime(1900, 1, 1),
            row["entity_ref"],
            row["task_name"],
        ),
        reverse=True,
    )

    selected_resource_type_name = (
        resource_type_name_map.get(
            selected_resource_type_id,
            "",
        )
        if selected_resource_type_id is not None
        else ""
    )

    selected_resource_name = (
        resource_name_map.get(
            selected_resource_id,
            "",
        )
        if selected_resource_id is not None
        else ""
    )

    html = """
    <!doctype html>
    <html>
    <head>
        <meta charset="utf-8">
        <meta http-equiv="refresh" content="300">
        <title>Manufacturing Floor Display</title>

        <style>
            * {
                box-sizing: border-box;
            }

            body {
                margin: 0;
                background: #0f172a;
                color: #f8fafc;
                font-family: 'Segoe UI', Arial, sans-serif;
            }

            .header {
                background: #1e293b;
                border-bottom: 3px solid #334155;
                padding: 18px 28px;
                display: flex;
                justify-content: space-between;
                align-items: center;
                gap: 20px;
            }

            .header-title {
                font-size: 32px;
                font-weight: 800;
                letter-spacing: .4px;
            }

            .header-date {
                color: #94a3b8;
                font-size: 17px;
                white-space: nowrap;
            }

            .filters {
                background: #1e293b;
                border-bottom: 1px solid #334155;
                padding: 14px 28px 18px;
            }

            .filter-form {
                display: flex;
                align-items: end;
                gap: 16px;
                flex-wrap: wrap;
            }

            .filter-group {
                display: flex;
                flex-direction: column;
                gap: 6px;
            }

            .filter-group label {
                color: #94a3b8;
                font-size: 13px;
                font-weight: 700;
                text-transform: uppercase;
                letter-spacing: .5px;
            }

            .filter-select {
                min-width: 250px;
                padding: 9px 12px;
                border: 1px solid #475569;
                border-radius: 7px;
                background: #0f172a;
                color: #f8fafc;
                font-size: 16px;
            }

            .filter-select:disabled {
                color: #64748b;
                cursor: not-allowed;
            }

            .clear-btn {
                display: inline-block;
                padding: 10px 16px;
                border: 1px solid #64748b;
                border-radius: 7px;
                color: #cbd5e1;
                text-decoration: none;
                font-size: 15px;
            }

            .filter-summary {
                margin-top: 11px;
                color: #94a3b8;
                font-size: 14px;
            }

            .board {
                padding: 22px 24px 30px;
                display: grid;
                grid-template-columns: repeat(3, minmax(0, 1fr));
                gap: 18px;
                align-items: start;
            }

            .column {
                min-width: 0;
            }

            .column-header {
                padding: 13px 14px;
                border-radius: 9px 9px 0 0;
                font-size: 21px;
                font-weight: 800;
                display: flex;
                justify-content: space-between;
                align-items: center;
                gap: 10px;
            }

            .header-not-started {
                background: #374151;
                color: #f8fafc;
            }

            .header-in-process {
                background: #1d4ed8;
                color: #eff6ff;
            }

            .header-completed {
                background: #166534;
                color: #f0fdf4;
            }

            .count {
                font-size: 16px;
                opacity: .8;
                white-space: nowrap;
            }

            .column-body {
                background: #1e293b;
                border-radius: 0 0 9px 9px;
                min-height: 360px;
                padding: 10px;
                display: flex;
                flex-direction: column;
                gap: 10px;
            }

            .task-card {
                background: #111827;
                border: 1px solid #334155;
                border-radius: 8px;
                padding: 13px 14px;
            }

            .task-name {
                font-size: 20px;
                font-weight: 800;
                line-height: 1.2;
                color: #f8fafc;
                margin-bottom: 7px;
            }

            .task-ref {
                font-size: 17px;
                font-weight: 700;
                color: #93c5fd;
            }

            .customer {
                color: #cbd5e1;
                font-size: 14px;
                margin-top: 2px;
            }

            .meta-grid {
                margin-top: 9px;
                display: grid;
                grid-template-columns: 1fr 1fr;
                gap: 5px 10px;
                font-size: 13px;
            }

            .meta-label {
                color: #64748b;
            }

            .meta-value {
                color: #cbd5e1;
                font-weight: 600;
                overflow-wrap: anywhere;
            }

            .active {
                margin-top: 10px;
                padding: 6px 8px;
                border-radius: 6px;
                background: #78350f;
                color: #fde68a;
                font-size: 13px;
                font-weight: 800;
            }

            .empty {
                color: #64748b;
                padding: 34px 10px;
                text-align: center;
                font-size: 17px;
            }

            @media (max-width: 1100px) {
                .board {
                    grid-template-columns: 1fr;
                }

                .column-body {
                    min-height: 100px;
                }
            }
        </style>
    </head>

    <body>
        <div class="header">
            <div class="header-title">
                FSG Signs Manufacturing Floor
            </div>

            <div class="header-date">
                {{ now_str }}
            </div>
        </div>

        <div class="filters">
            <form
                class="filter-form"
                method="get"
                action="{{ url_for('manufacturing_floor_display') }}"
            >
                <div class="filter-group">
                    <label for="resource_type">
                        Resource Type
                    </label>

                    <select
                        class="filter-select"
                        id="resource_type"
                        name="resource_type"
                        onchange="
                            document.getElementById('resource').value='';
                            this.form.submit();
                        "
                    >
                        <option value="">
                            All Resource Types
                        </option>

                        {% for rt in resource_types %}
                        <option
                            value="{{ rt.resource_type_id }}"
                            {% if selected_resource_type_id == rt.resource_type_id %}
                                selected
                            {% endif %}
                        >
                            {{ rt.resource_description }}
                        </option>
                        {% endfor %}
                    </select>
                </div>

                <div class="filter-group">
                    <label for="resource">
                        Resource Name
                    </label>

                    <select
                        class="filter-select"
                        id="resource"
                        name="resource"
                        onchange="this.form.submit();"
                        {% if selected_resource_type_id is none %}
                            disabled
                        {% endif %}
                    >
                        <option value="">
                            All Resources
                        </option>

                        {% for r in resource_options %}
                        <option
                            value="{{ r.resource_id }}"
                            {% if selected_resource_id == r.resource_id %}
                                selected
                            {% endif %}
                        >
                            {{ r.resource_name }}
                        </option>
                        {% endfor %}
                    </select>
                </div>

                {% if
                    selected_resource_type_id is not none
                    or selected_resource_id is not none
                %}
                <a
                    class="clear-btn"
                    href="{{ url_for('manufacturing_floor_display') }}"
                >
                    Clear Filters
                </a>
                {% endif %}
            </form>

            <div class="filter-summary">
                Showing:

                {% if selected_resource_type_name %}
                    {{ selected_resource_type_name }}
                {% else %}
                    All Resource Types
                {% endif %}

                {% if selected_resource_name %}
                    &nbsp;&rsaquo;&nbsp;
                    {{ selected_resource_name }}
                {% endif %}
            </div>
        </div>

        <div class="board">
            <section class="column">
                <div class="column-header header-not-started">
                    <span>
                        Tasks Not Started
                    </span>

                    <span class="count">
                        {{ columns.not_started|length }}
                    </span>
                </div>

                <div class="column-body">
                    {% for task in columns.not_started %}
                    <div class="task-card">
                        <div class="task-name">
                            {{ task.task_name }}
                        </div>

                        <div class="task-ref">
                            {{ task.entity_ref }}
                        </div>

                        {% if task.customer %}
                        <div class="customer">
                            {{ task.customer }}
                        </div>
                        {% endif %}

                        <div class="meta-grid">
                            <div class="meta-label">
                                Resource Type
                            </div>

                            <div class="meta-value">
                                {{ task.resource_type }}
                            </div>

                            <div class="meta-label">
                                Resource
                            </div>

                            <div class="meta-value">
                                {{ task.resource_name }}
                            </div>

                            <div class="meta-label">
                                Due
                            </div>

                            <div class="meta-value">
                                {{
                                    task.due_date.strftime('%b %d')
                                    if task.due_date
                                    else '—'
                                }}
                            </div>

                            <div class="meta-label">
                                Planned / Logged
                            </div>

                            <div class="meta-value">
                                {{ fmt_hours(task.planned_hours) }}
                                /
                                {{ fmt_mins(task.logged_minutes) }}
                            </div>
                        </div>
                    </div>

                    {% else %}
                    <div class="empty">
                        No tasks not started.
                    </div>
                    {% endfor %}
                </div>
            </section>

            <section class="column">
                <div class="column-header header-in-process">
                    <span>
                        Tasks In Process
                    </span>

                    <span class="count">
                        {{ columns.in_process|length }}
                    </span>
                </div>

                <div class="column-body">
                    {% for task in columns.in_process %}
                    <div class="task-card">
                        <div class="task-name">
                            {{ task.task_name }}
                        </div>

                        <div class="task-ref">
                            {{ task.entity_ref }}
                        </div>

                        {% if task.customer %}
                        <div class="customer">
                            {{ task.customer }}
                        </div>
                        {% endif %}

                        <div class="meta-grid">
                            <div class="meta-label">
                                Resource Type
                            </div>

                            <div class="meta-value">
                                {{ task.resource_type }}
                            </div>

                            <div class="meta-label">
                                Resource
                            </div>

                            <div class="meta-value">
                                {{ task.resource_name }}
                            </div>

                            <div class="meta-label">
                                Due
                            </div>

                            <div class="meta-value">
                                {{
                                    task.due_date.strftime('%b %d')
                                    if task.due_date
                                    else '—'
                                }}
                            </div>

                            <div class="meta-label">
                                Planned / Logged
                            </div>

                            <div class="meta-value">
                                {{ fmt_hours(task.planned_hours) }}
                                /
                                {{ fmt_mins(task.logged_minutes) }}
                            </div>
                        </div>

                        {% if task.active_names %}
                        <div class="active">
                            &#9733; Active:
                            {{ task.active_names|join(', ') }}
                        </div>
                        {% endif %}
                    </div>

                    {% else %}
                    <div class="empty">
                        No tasks currently in process.
                    </div>
                    {% endfor %}
                </div>
            </section>

            <section class="column">
                <div class="column-header header-completed">
                    <span>
                        Tasks Completed Today
                    </span>

                    <span class="count">
                        {{ columns.completed_today|length }}
                    </span>
                </div>

                <div class="column-body">
                    {% for task in columns.completed_today %}
                    <div class="task-card">
                        <div class="task-name">
                            {{ task.task_name }}
                        </div>

                        <div class="task-ref">
                            {{ task.entity_ref }}
                        </div>

                        {% if task.customer %}
                        <div class="customer">
                            {{ task.customer }}
                        </div>
                        {% endif %}

                        <div class="meta-grid">
                            <div class="meta-label">
                                Resource Type
                            </div>

                            <div class="meta-value">
                                {{ task.resource_type }}
                            </div>

                            <div class="meta-label">
                                Resource
                            </div>

                            <div class="meta-value">
                                {{ task.resource_name }}
                            </div>

                            <div class="meta-label">
                                Completed
                            </div>

                            <div class="meta-value">
                                {{
                                    task.actual_finish_date.strftime('%I:%M %p')
                                    if task.actual_finish_date
                                    else '—'
                                }}
                            </div>

                            <div class="meta-label">
                                Planned / Logged
                            </div>

                            <div class="meta-value">
                                {{ fmt_hours(task.planned_hours) }}
                                /
                                {{ fmt_mins(task.logged_minutes) }}
                            </div>
                        </div>
                    </div>

                    {% else %}
                    <div class="empty">
                        No tasks completed today.
                    </div>
                    {% endfor %}
                </div>
            </section>
        </div>
    </body>
    </html>
    """

    return render_template_string(
        html,
        columns=columns,
        resource_types=resource_types,
        resource_options=resource_options,
        selected_resource_type_id=selected_resource_type_id,
        selected_resource_id=selected_resource_id,
        selected_resource_type_name=selected_resource_type_name,
        selected_resource_name=selected_resource_name,
        now_str=now.strftime(
            "%A, %B %d, %Y — %I:%M %p"
        ),
        fmt_mins=fmt_mins,
        fmt_hours=fmt_hours,
    )
########################################################################################################################
########################################################################################################################
########################################################################################################################





########################################################################################################################
########################################################################################################################
########################################################################################################################
"""QC/SURVEY RESPONSES"""
########################################################################################################################
########################################################################################################################
########################################################################################################################
########################################################################################################################
@app.route("/qc-survey-responses")
@login_required
@office_required
def qc_survey_responses():

    from types import SimpleNamespace

    search_type = request.args.get(
        "search_type",
        "job",
    ).strip().lower()

    search_value = request.args.get(
        "q",
        "",
    ).strip()

    results = []

    # =========================================================================
    # ONLY SEARCH WHEN USER ENTERS SOMETHING
    # =========================================================================

    if search_value:

        try:
            search_id = int(search_value)

        except (TypeError, ValueError):
            search_id = None

        if search_id is None:

            flash(
                "Enter a valid Job Number or Task ID.",
                "error",
            )

        else:

            # =================================================================
            # FIND TASK IDS
            # =================================================================

            task_ids = []

            # -------------------------------------------------------------
            # SEARCH BY TASK ID
            # -------------------------------------------------------------

            if search_type == "task":

                task = fetch_one(
                    """
                    SELECT task_id
                    FROM tasks
                    WHERE task_id = ?
                    """,
                    [search_id],
                )

                if task:
                    task_ids.append(
                        task.task_id
                    )

            # -------------------------------------------------------------
            # SEARCH BY JOB NUMBER
            # -------------------------------------------------------------

            else:

                line_items = fetch_all(
                    """
                    SELECT
                        line_item_id
                    FROM line_items
                    WHERE job_number_id = ?
                    ORDER BY line_item_id
                    """,
                    [search_id],
                )

                for line_item in line_items:

                    cost_codes = fetch_all(
                        """
                        SELECT
                            cost_code_id
                        FROM cost_codes
                        WHERE line_item_id = ?
                        ORDER BY cost_code_id
                        """,
                        [
                            line_item.line_item_id
                        ],
                    )

                    for cost_code in cost_codes:

                        tasks = fetch_all(
                            """
                            SELECT
                                task_id
                            FROM tasks
                            WHERE cost_code_id = ?
                            ORDER BY task_id
                            """,
                            [
                                cost_code.cost_code_id
                            ],
                        )

                        for task in tasks:

                            task_ids.append(
                                task.task_id
                            )

            # Remove duplicates while maintaining order.
            task_ids = list(
                dict.fromkeys(task_ids)
            )

            # =================================================================
            # LOAD EACH TASK THAT HAS QC / SURVEY ANSWERS
            # =================================================================

            for current_task_id in task_ids:

                answers = fetch_all(
                    """
                    SELECT
                        qc_answer_id,
                        qc_question_id,
                        answer_value,
                        employee_id,
                        submitted_date
                    FROM qc_form_answers
                    WHERE task_id = ?
                    ORDER BY
                        submitted_date,
                        qc_answer_id
                    """,
                    [current_task_id],
                )

                # No responses for this task.
                if not answers:
                    continue

                # =============================================================
                # LOAD TASK
                # =============================================================

                task = fetch_one(
                    """
                    SELECT
                        task_id,
                        task_name,
                        task_description,
                        cost_code_id,
                        resource_type_id,
                        survey_type
                    FROM tasks
                    WHERE task_id = ?
                    """,
                    [current_task_id],
                )

                if not task:
                    continue

                # =============================================================
                # COST CODE
                # =============================================================

                cost_code = None

                if task.cost_code_id:

                    cost_code = fetch_one(
                        """
                        SELECT
                            cost_code_id,
                            coins_cost_code,
                            cost_code_description,
                            line_item_id
                        FROM cost_codes
                        WHERE cost_code_id = ?
                        """,
                        [task.cost_code_id],
                    )

                # =============================================================
                # LINE ITEM / JOB
                # =============================================================

                line_item = None

                if (
                    cost_code
                    and cost_code.line_item_id
                ):

                    line_item = fetch_one(
                        """
                        SELECT
                            line_item_id,
                            line_item_description,
                            job_number_id,
                            work_order_id
                        FROM line_items
                        WHERE line_item_id = ?
                        """,
                        [
                            cost_code.line_item_id
                        ],
                    )

                # =============================================================
                # BUILD RESPONSE GROUPS
                #
                # One submitted form generally has the same submitted_date
                # and employee_id.
                # =============================================================

                submissions = {}

                for answer in answers:

                    submission_key = (
                        answer.submitted_date,
                        answer.employee_id,
                    )

                    if submission_key not in submissions:

                        # -----------------------------------------------------
                        # EMPLOYEE
                        # -----------------------------------------------------

                        employee_name = ""

                        if answer.employee_id:

                            answer_employee = fetch_one(
                                """
                                SELECT
                                    employee_first_name,
                                    employee_last_name
                                FROM employees
                                WHERE employee_id = ?
                                """,
                                [answer.employee_id],
                            )

                            if answer_employee:

                                employee_name = (
                                    f"{answer_employee.employee_first_name or ''} "
                                    f"{answer_employee.employee_last_name or ''}"
                                ).strip()

                        submissions[
                            submission_key
                        ] = {
                            "submitted_date":
                                answer.submitted_date,

                            "employee_name":
                                employee_name,

                            "answers": [],

                            "form_names": [],
                        }

                    submission = submissions[
                        submission_key
                    ]

                    # ---------------------------------------------------------
                    # ADDITIONAL NOTES
                    # ---------------------------------------------------------

                    if (
                        answer.answer_value
                        and str(
                            answer.answer_value
                        ).startswith("NOTES:")
                    ):

                        submission[
                            "answers"
                        ].append(
                            {
                                "section_name":
                                    "",

                                "question_text":
                                    "Additional Notes",

                                "answer_value":
                                    str(
                                        answer.answer_value
                                    )[6:],

                                "sort_order":
                                    999999,
                            }
                        )

                        continue

                    # ---------------------------------------------------------
                    # QUESTION INFORMATION
                    # ---------------------------------------------------------

                    question = None

                    if answer.qc_question_id:

                        question = fetch_one(
                            """
                            SELECT
                                qc_question_id,
                                qc_form_id,
                                question_text,
                                section_name,
                                sort_order
                            FROM qc_form_questions
                            WHERE qc_question_id = ?
                            """,
                            [
                                answer.qc_question_id
                            ],
                        )

                    if question:

                        form_name = ""

                        if question.qc_form_id:

                            form = fetch_one(
                                """
                                SELECT
                                    form_name
                                FROM qc_forms
                                WHERE qc_form_id = ?
                                """,
                                [
                                    question.qc_form_id
                                ],
                            )

                            if form:

                                form_name = (
                                    form.form_name
                                    or ""
                                )

                                if (
                                    form_name
                                    and form_name
                                    not in submission[
                                        "form_names"
                                    ]
                                ):

                                    submission[
                                        "form_names"
                                    ].append(
                                        form_name
                                    )

                        submission[
                            "answers"
                        ].append(
                            {
                                "section_name":
                                    (
                                        question.section_name
                                        or ""
                                    ),

                                "question_text":
                                    (
                                        question.question_text
                                        or ""
                                    ),

                                "answer_value":
                                    (
                                        answer.answer_value
                                        or ""
                                    ),

                                "sort_order":
                                    (
                                        question.sort_order
                                        or 0
                                    ),
                            }
                        )

                # =============================================================
                # CONVERT SUBMISSIONS TO DISPLAY OBJECTS
                # =============================================================

                submission_rows = []

                for submission in submissions.values():

                    submission[
                        "answers"
                    ].sort(
                        key=lambda x: (
                            x["sort_order"],
                            x["question_text"],
                        )
                    )

                    form_name = "QC / Survey Form"

                    if submission[
                        "form_names"
                    ]:

                        form_name = ", ".join(
                            submission[
                                "form_names"
                            ]
                        )

                    elif task.survey_type:

                        survey_form = fetch_one(
                            """
                            SELECT
                                form_name
                            FROM qc_forms
                            WHERE qc_form_id = ?
                            """,
                            [task.survey_type],
                        )

                        if survey_form:

                            form_name = (
                                survey_form.form_name
                                or form_name
                            )

                    submission_rows.append(
                        SimpleNamespace(
                            submitted_date=
                                submission[
                                    "submitted_date"
                                ],

                            employee_name=
                                submission[
                                    "employee_name"
                                ],

                            form_name=form_name,

                            answers=
                                submission[
                                    "answers"
                                ],
                        )
                    )

                submission_rows.sort(
                    key=lambda x: (
                        x.submitted_date
                        or datetime.min
                    ),
                    reverse=True,
                )

                results.append(
                    SimpleNamespace(

                        task_id=
                            task.task_id,

                        task_name=
                            (
                                task.task_name
                                or ""
                            ),

                        task_description=
                            (
                                task.task_description
                                or ""
                            ),

                        job_number_id=(
                            line_item.job_number_id
                            if line_item
                            else None
                        ),

                        work_order_id=(
                            line_item.work_order_id
                            if line_item
                            else None
                        ),

                        line_item_description=(
                            line_item.line_item_description
                            if line_item
                            else ""
                        ),

                        cost_code=(
                            cost_code.coins_cost_code
                            if cost_code
                            else ""
                        ),

                        submissions=
                            submission_rows,
                    )
                )

    # =========================================================================
    # SCREEN
    # =========================================================================

    html = """
    <!doctype html>

    <html>

    <head>

        <title>
            QC and Survey Responses
        </title>

        <style>

            * {
                box-sizing: border-box;
            }

            body {
                font-family: Arial, sans-serif;
                margin: 0;
                padding: 30px;
                background: #f5f6f8;
                color: #212529;
            }

            .wrap {
                max-width: 1300px;
                margin: auto;
            }

            .card {
                background: white;
                padding: 26px;
                border-radius: 10px;
                box-shadow:
                    0 2px 10px
                    rgba(0,0,0,.08);
                margin-bottom: 20px;
            }

            .btn {
                display: inline-block;
                padding: 11px 18px;
                border: none;
                border-radius: 7px;
                background: #1f4e79;
                color: white;
                text-decoration: none;
                font-weight: bold;
                cursor: pointer;
            }

            .gray {
                background: #6c757d;
            }

            .search-row {
                display: grid;
                grid-template-columns:
                    190px
                    minmax(250px, 1fr)
                    auto;
                gap: 10px;
                align-items: end;
            }

            label {
                display: block;
                font-weight: bold;
                margin-bottom: 5px;
            }

            input,
            select {
                width: 100%;
                padding: 11px;
                border: 1px solid #ccc;
                border-radius: 6px;
                font-size: 15px;
            }

            h1 {
                margin-top: 0;
                color: #1f4e79;
            }

            h2 {
                color: #1f4e79;
                margin-top: 0;
            }

            .task-header {
                display: flex;
                justify-content:
                    space-between;
                gap: 20px;
                align-items: start;
                flex-wrap: wrap;
            }

            .task-meta {
                color: #666;
                margin-top: 5px;
                line-height: 1.5;
            }

            .submission {
                margin-top: 20px;
                border: 1px solid #ccc;
                border-radius: 8px;
                overflow: hidden;
            }

            .submission-header {
                padding: 14px 16px;
                background: #1f4e79;
                color: white;
            }

            .submission-title {
                font-size: 18px;
                font-weight: bold;
            }

            .submission-meta {
                margin-top: 4px;
                font-size: 13px;
                opacity: .85;
            }

            .answers {
                padding: 0 16px 16px;
            }

            .section-title {
                color: #1f4e79;
                font-weight: bold;
                font-size: 17px;
                margin-top: 22px;
                padding-bottom: 5px;
                border-bottom:
                    2px solid #1f4e79;
            }

            .answer-row {
                display: grid;
                grid-template-columns:
                    minmax(250px, 45%)
                    1fr;
                gap: 18px;
                padding: 10px 0;
                border-bottom:
                    1px solid #eee;
            }

            .question {
                font-weight: bold;
            }

            .answer {
                white-space: pre-wrap;
            }

            .no-results {
                color: #666;
                font-size: 16px;
            }

            .flash-error {
                background: #f8d7da;
                color: #842029;
                padding: 10px;
                border-radius: 6px;
                margin-bottom: 12px;
            }

            @media(max-width: 700px) {

                body {
                    padding: 12px;
                }

                .search-row {
                    grid-template-columns: 1fr;
                }

                .answer-row {
                    grid-template-columns: 1fr;
                    gap: 4px;
                }

            }

        </style>

    </head>


    <body>

        <div class="wrap">

            <a
                class="btn gray"
                href="{{ url_for('home') }}"
                style="margin-bottom:20px;"
            >
                ← Home
            </a>


            {% with messages =
                get_flashed_messages(
                    with_categories=true
                )
            %}

                {% if messages %}

                    {% for category, message
                        in messages
                    %}

                        <div
                            class="flash-{{ category }}"
                        >
                            {{ message }}
                        </div>

                    {% endfor %}

                {% endif %}

            {% endwith %}


            <div class="card">

                <h1>
                    QC and Survey Responses
                </h1>

                <p>
                    Search submitted QC and survey
                    forms by Job Number or Task ID.
                </p>


                <form
                    method="get"
                    action="{{
                        url_for(
                            'qc_survey_responses'
                        )
                    }}"
                >

                    <div class="search-row">

                        <div>

                            <label>
                                Search By
                            </label>

                            <select
                                name="search_type"
                            >

                                <option
                                    value="job"
                                    {% if search_type == 'job' %}
                                        selected
                                    {% endif %}
                                >
                                    Job Number
                                </option>

                                <option
                                    value="task"
                                    {% if search_type == 'task' %}
                                        selected
                                    {% endif %}
                                >
                                    Task ID
                                </option>

                            </select>

                        </div>


                        <div>

                            <label>
                                Job Number / Task ID
                            </label>

                            <input
                                type="text"
                                name="q"
                                value="{{ search_value }}"
                                placeholder="Enter number"
                                required
                            >

                        </div>


                        <div>

                            <button
                                class="btn"
                                type="submit"
                            >
                                Search
                            </button>

                        </div>

                    </div>

                </form>

            </div>


            {% if search_value %}

                {% if results %}

                    {% for task in results %}

                        <div class="card">

                            <div class="task-header">

                                <div>

                                    <h2>
                                        Task {{ task.task_id }}
                                        —
                                        {{
                                            task.task_name
                                            or 'Unnamed Task'
                                        }}
                                    </h2>

                                    <div class="task-meta">

                                        {% if task.job_number_id %}

                                            <strong>
                                                Job:
                                            </strong>

                                            {{ task.job_number_id }}

                                        {% elif task.work_order_id %}

                                            <strong>
                                                WO:
                                            </strong>

                                            {{ task.work_order_id }}

                                        {% endif %}


                                        {% if task.line_item_description %}

                                            <br>

                                            <strong>
                                                Line Item:
                                            </strong>

                                            {{
                                                task.line_item_description
                                            }}

                                        {% endif %}


                                        {% if task.cost_code %}

                                            <br>

                                            <strong>
                                                Cost Code:
                                            </strong>

                                            {{ task.cost_code }}

                                        {% endif %}

                                    </div>

                                </div>


                                <a
                                    class="btn"
                                    href="{{
                                        url_for(
                                            'office_task_detail',
                                            task_id=
                                                task.task_id
                                        )
                                    }}"
                                >
                                    View Task
                                </a>

                            </div>


                            {% for submission
                                in task.submissions
                            %}

                                <div class="submission">

                                    <div
                                        class="submission-header"
                                    >

                                        <div
                                            class="submission-title"
                                        >
                                            {{
                                                submission.form_name
                                            }}
                                        </div>

                                        <div
                                            class="submission-meta"
                                        >

                                            Submitted by

                                            {{
                                                submission.employee_name
                                                or 'Unknown'
                                            }}

                                            {% if
                                                submission.submitted_date
                                            %}

                                                on

                                                {{
                                                    submission.submitted_date.strftime(
                                                        '%b %d, %Y %I:%M %p'
                                                    )
                                                }}

                                            {% endif %}

                                        </div>

                                    </div>


                                    <div class="answers">

                                        {% set ns =
                                            namespace(
                                                section=''
                                            )
                                        %}

                                        {% for answer
                                            in submission.answers
                                        %}

                                            {% if
                                                answer.section_name
                                                and
                                                answer.section_name
                                                != ns.section
                                            %}

                                                <div
                                                    class="section-title"
                                                >
                                                    {{
                                                        answer.section_name
                                                    }}
                                                </div>

                                                {% set ns.section =
                                                    answer.section_name
                                                %}

                                            {% endif %}


                                            <div
                                                class="answer-row"
                                            >

                                                <div
                                                    class="question"
                                                >
                                                    {{
                                                        answer.question_text
                                                    }}
                                                </div>

                                                <div
                                                    class="answer"
                                                >
                                                    {{
                                                        answer.answer_value
                                                        or '—'
                                                    }}
                                                </div>

                                            </div>

                                        {% endfor %}

                                    </div>

                                </div>

                            {% endfor %}

                        </div>

                    {% endfor %}

                {% else %}

                    <div class="card no-results">

                        No QC or survey responses
                        were found for this search.

                    </div>

                {% endif %}

            {% endif %}

        </div>

    </body>

    </html>
    """

    return render_template_string(
        html,
        search_type=search_type,
        search_value=search_value,
        results=results,
    )
########################################################################################################################
########################################################################################################################
########################################################################################################################
########################################################################################################################



if __name__ == "__main__":
    app.run(debug=True)
