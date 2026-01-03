import os
import random
import streamlit as st
from openpyxl import load_workbook


# --- 1. DEFINITIONS ---

def parse_campaign_data():
    """Parses 'campaign_data.xlsx' for missions, intensity, and duration."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(script_dir, 'campaign_data.xlsx')

    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    wb = load_workbook(file_path, data_only=True)
    ws = wb.active

    campaign_data = {}
    for col in ws.iter_cols(values_only=True):
        campaign_name = col[0]
        if campaign_name is None:
            continue

        campaign_name = str(campaign_name).strip()
        campaign_dict = {}
        current_key = None

        for cell_value in col[1:]:
            if cell_value is None:
                current_key = None
                continue
            if current_key is None:
                current_key = str(cell_value).strip().lower()
                campaign_dict[current_key] = []
            else:
                val = cell_value
                if current_key == "duration" and isinstance(val, (int, float)):
                    val = int(val)
                campaign_dict[current_key].append(val)
        campaign_data[campaign_name] = campaign_dict
    return campaign_data


def parse_intensity_data():
    """
    Parses 'Intensity calcs.xlsx' into a structured dictionary.
    Updated to include Scale_min (column 5) and Scale_max (column 6).
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(script_dir, 'Intensity calcs.xlsx')

    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    wb = load_workbook(file_path, data_only=True)
    ws = wb.active

    intensity_lookup = {}

    # Iterate through rows, skipping header
    for row in ws.iter_rows(min_row=2, values_only=True):
        name = str(row[0]).strip()
        intensity_lookup[name] = {
            "description": row[1],
            "probability": row[2],
            "payout": row[3],
            "scale_min": row[4],  # New: Scale_min multiplier
            "scale_max": row[5]  # New: Scale_max multiplier
        }

    return intensity_lookup

def calculate_pay_split(total_pay, duration):
    """
    Calculates a split between a monthly retainer and a one-time bonus.
    Ensures both are multiples of 5 and at least 25% of total value.
    """
    possible_splits = []

    # Iterate through possible monthly retainer values (increments of 5)
    # The max possible retainer * duration cannot exceed 75% of total (to leave 25% for bonus)
    for m in range(0, (total_pay // duration) + 5, 5):
        total_retainer = m * duration
        bonus = total_pay - total_retainer

        # Constraints:
        # 1. Total Retainer >= 25%
        # 2. Bonus >= 25%
        # 3. Bonus is a multiple of 5 (m is already guaranteed by range step)
        if total_retainer >= 0.25 * total_pay and bonus >= 0.25 * total_pay:
            if bonus % 5 == 0:
                possible_splits.append((m, bonus))

    if possible_splits:
        return random.choice(possible_splits)

    # Fallback: simple 50/50 split rounded to nearest 5 if no perfect matches found
    m_raw = (total_pay // 2) // duration
    m_fallback = 5 * round(m_raw / 5)
    bonus_fallback = total_pay - (m_fallback * duration)
    return m_fallback, bonus_fallback


def calculate_scale_pv(intensity_name, intensity_data):
    """
    Calculates a digestible PV value based on intensity scale limits.
    Base PV is 120. Result is rounded to nearest 5.
    """
    stats = intensity_data.get(intensity_name.strip(), {})
    scale_min = stats.get('scale_min', 1.0)
    scale_max = stats.get('scale_max', 1.0)

    # Pick a random multiplier between the min and max for this intensity
    multiplier = random.uniform(scale_min, scale_max)
    raw_pv = 120 * multiplier

    # Round to the nearest multiple of 5 to make it "digestible"
    return int(5 * round(raw_pv / 5))


def get_monthly_mission_count(intensity_name):
    """
    Determines how many missions occur in a single month based on intensity.
    Logic:
    - Medium (Avg 1): 0 (20%), 1 (60%), 2 (20%)
    - High (Avg 2): 1 (20%), 2 (60%), 3 (20%)
    - etc.
    """
    # Define the weights for each intensity
    # Format: { 'Intensity': [list_of_possible_counts], [weights_for_each] }
    intensity_rules = {
        "Very low": {"counts": [0, 1], "weights": [75, 25]},
        "Low": {"counts": [0, 1], "weights": [50, 50]},
        "Medium": {"counts": [0, 1, 2], "weights": [20, 60, 20]},
        "High": {"counts": [1, 2, 3], "weights": [20, 60, 20]},
        "Very high": {"counts": [2, 3, 4], "weights": [20, 60, 20]}
    }

    rule = intensity_rules.get(intensity_name.strip(), intensity_rules["Medium"])

    # random.choices returns a list, so we take the first [0] element
    return random.choices(rule["counts"], weights=rule["weights"])[0]


def generate_mission_schedule(duration, intensity_name, campaign_type, campaign_data):
    """
    Creates a month-by-month list of mission counts and types.
    Ensures the same mission never appears twice in a row.
    """
    schedule = []
    # Get the list of allowed missions for this specific campaign type
    available_missions = campaign_data.get(campaign_type, {}).get('missions', ["Standard Combat"])

    last_mission = None  # Track the previous mission to avoid repeats

    for month in range(1, duration + 1):
        count = get_monthly_mission_count(intensity_name)

        month_missions = []
        for _ in range(count):
            # Create a list of choices excluding the last mission picked
            valid_choices = [m for m in available_missions if m != last_mission]

            # If for some reason only one mission type exists in the Excel,
            # fallback to it to avoid a crash
            if not valid_choices:
                valid_choices = available_missions

            current_mission = random.choice(valid_choices)
            month_missions.append(current_mission)
            last_mission = current_mission  # Update the tracker

        schedule.append({
            "month": month,
            "count": count,
            "types": month_missions
        })
    return schedule


def parse_contract_parameters():
    """
    Parses 'contract_parameters.xlsx' to get lists of actors and salvage terms.
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(script_dir, 'contract_parameters.xlsx')

    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    wb = load_workbook(file_path, data_only=True)
    ws = wb.active

    actors = []
    salvage_terms = []

    # Iterate through the rows, skipping header
    for row in ws.iter_rows(min_row=2, values_only=True):
        # row[0] is Employer/Actor, row[1] is Salvage
        if row[0]:
            actors.append(str(row[0]).strip())
        if row[1]:
            salvage_terms.append(str(row[1]).strip())

    return {"actors": actors, "salvage": salvage_terms}


def generate_random_campaign(campaign_data, intensity_data, contract_params):
    # 1. Pick random campaign and core details
    campaign_type = random.choice(list(campaign_data.keys()))
    details = campaign_data[campaign_type]

    intensity_name = random.choice(details.get('intensity', ["Standard"]))
    duration = random.choice(details.get('duration', [1]))

    # 2. Pick Employer, Opponent, and Salvage
    employer = random.choice(contract_params["actors"])
    remaining_actors = [a for a in contract_params["actors"] if a != employer]
    opponent = random.choice(remaining_actors)
    salvage = random.choice(contract_params["salvage"])

    # 3. Get intensity stats
    stats = intensity_data.get(intensity_name.strip(), {})
    intensity_description = stats.get('description', "No description available")
    base_monthly_payout = stats.get('payout', 0)

    # 4. Calculate Scale PV
    pv_value = calculate_scale_pv(intensity_name, intensity_data)

    # 5. Adjust Financials by PV and 20% Variance
    base_total_pay = duration * base_monthly_payout
    pv_adjusted_pay = (pv_value / 120) * base_total_pay

    # NEW: Introduce 20% variance (random factor between 0.8 and 1.2)
    variance_factor = random.uniform(0.8, 1.2)
    final_raw_pay = pv_adjusted_pay * variance_factor

    # Ensure the total is "digestible" (multiple of 5) before splitting
    digestible_total_pay = int(5 * round(final_raw_pay / 5))

    # Split the adjusted total into retainer and bonus
    monthly_retainer, bonus = calculate_pay_split(digestible_total_pay, duration)

    # 6. Generate the Mission Schedule
    mission_schedule = generate_mission_schedule(duration, intensity_name, campaign_type, campaign_data)

    # 7. Format outputs
    duration_text = f"{duration} month" if duration == 1 else f"{duration} months"

    output = [
        f"Employer: {employer}",
        f"Opponent: {opponent}",
        f"Campaign: {campaign_type}",
        f"Intensity: {intensity_name} ({intensity_description})",
        f"Duration: {duration_text}",
        f"Salvage: {salvage}",
        f"Pay (retainer): {monthly_retainer} per month",
        f"Pay (bonus): {bonus} one-time",
        f"Scale: Lance (minimum of {pv_value} PV)",
        "\nMission Schedule:"
    ]

    for item in mission_schedule:
        m_text = "mission" if item["count"] == 1 else "missions"
        line = f"Month {item['month']}: {item['count']} {m_text}"
        if item["count"] > 0:
            line += f" ({', '.join(item['types'])})"
        output.append(line)

    return "\n".join(output)

## --- EXECUTION ---
st.title("Campaign Generator")

if st.button("Generate New Campaign"):
    all_campaigns = parse_campaign_data()
    intensity_stats = parse_intensity_data()
    contract_params = parse_contract_parameters()

    random_setup = generate_random_campaign(all_campaigns, intensity_stats, contract_params)

    # This displays the result in a nice box on the website
    st.text_area("Result", value=random_setup, height=400)