import pandas as pd
from gurobipy import Model, GRB, LinExpr
import math
import copy
import time
from collections import defaultdict

def load_data(excel_file):
    """
    Load and preprocess all data from the Excel file.
    
    Args:
        excel_file (str): Path to the Excel file containing the scheduling data.
        
    Returns:
        tuple: Contains dataframes and mappings needed for the model.
    """
    print("Loading data...")
    
    # Load and preprocess the assignments data
    df = pd.read_excel(excel_file, sheet_name='Assignments')
    df = df.rename(columns={df.columns[0]: 'Course'})
    df = df.rename(columns={df.columns[1]: 'Instructor'})
    df = df.rename(columns={df.columns[2]: 'Capacity'})
    df = df.rename(columns={df.columns[3]: '# Sections'})
    df = df.iloc[:, :-4]
    
    # Print summary and filter data
    print(f"Total sections: {df['# Sections'].sum()}")
    df = df.dropna(subset=['Course'])
    df['Course_Number'] = df['Course'].str.extract(r'(\d+)')
    
    # Load faculty data
    faculty_df = pd.read_excel(excel_file, sheet_name='Faculty')
    faculty_df = faculty_df.rename(columns={faculty_df.columns[0]: 'Instructor'})
    
    # Merge assignments with faculty data
    df = pd.merge(df, faculty_df, on='Instructor', how='left')
    
    # Load pre-scheduled courses
    df_pre_scheduled = pd.read_excel(excel_file, sheet_name='pre-scheduled')
    
    # Load constraints and preferences
    df_constraints = pd.read_excel(excel_file, sheet_name='Constraints & Preferences')
    
    # Load general preferences
    general_preferences_df = pd.read_excel(excel_file, sheet_name='General Preferences')
    general_preferences_df = general_preferences_df.rename(columns={
        general_preferences_df.columns[1]: 'Email', 
        general_preferences_df.columns[2]: 'Preference',
        general_preferences_df.columns[3]: 'Day Preference',
        general_preferences_df.columns[5]: 'Consecutive Preference'
    })
    
    # Create an aggregated dataframe (course/instructor/sections)
    aggregated_df = create_aggregated_dataframe(df)
    
    # Create section capacity map
    section_capacity_map = create_section_capacity_map(df)
    
    return df, aggregated_df, df_pre_scheduled, df_constraints, general_preferences_df, section_capacity_map

def create_section_capacity_map(df):
    """
    Creates a dictionary mapping (course, instructor, section_number) to capacity.

    Args:
        df (DataFrame): DataFrame containing course, instructor, and capacity information.

    Returns:
        dict: A dictionary where keys are (course, instructor, section_number)
              and values are the capacity of the sections.
    """
    # Sort by Course, Instructor, and Capacity to ensure correct section numbering
    df_sorted = df.sort_values(by=['Course', 'Instructor', 'Capacity'])

    section_capacity_map = {}

    # Group by (course, instructor) and assign section numbers
    for (course, instructor), group in df_sorted.groupby(['Course', 'Instructor']):
        section_number = 1
        for _, row in group.iterrows():
            num_sections = int(row['# Sections'])
            capacity = row['Capacity']
            course = course.strip()
            instructor = instructor.strip()

            # Assign section numbers based on the sorted order
            for _ in range(num_sections):
                section_capacity_map[(course, instructor, section_number)] = capacity
                section_number += 1

    return section_capacity_map

def create_aggregated_dataframe(df):
    """
    Group by Course, Instructor and sum the number of sections.
    
    Args:
        df (DataFrame): DataFrame containing course and instructor information.
        
    Returns:
        DataFrame: Aggregated DataFrame with total sections by course and instructor.
    """
    aggregated_df = df.groupby(['Course', 'Instructor', 'Course_Number', 'Email']).agg({
        '# Sections': 'sum'
    }).reset_index()
    
    return aggregated_df

def define_time_slots_and_days():
    """
    Define the available time slots and days for scheduling.
    
    Returns:
        tuple: Lists of time slots and days.
    """
    time_slots = [
        "8:30-10:00 AM",
        "10:00-11:30 AM",
        "11:30-1:00 PM",
        "1:00-2:30 PM",
        "2:30-4:00 PM",
        "4:00-5:30 PM",
        "6:00-7:30 PM",
        "7:30-9:00 PM"
    ]

    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
    
    return time_slots, days

def define_slot_percentages(days, time_slots):
    """
    Define the target percentage of classes for each day-time slot combination.
    
    Args:
        days (list): List of days.
        time_slots (list): List of time slots.
        
    Returns:
        dict: Dictionary mapping (day, time_slot) to percentage.
    """
    slot_percentages = {}
    
    # Default percentages
    default_percentage = 0.20
    wednesday_percentage = 0.25
    zero_percentage = 0.00
    
    # Assign percentage to each day-slot combination
    for day in days:
        for slot in time_slots:
            # Special cases
            if day == "Wednesday" and slot in ["8:30-10:00 AM", "10:00-11:30 AM", "11:30-1:00 PM", "1:00-2:30 PM"]:
                slot_percentages[(day, slot)] = wednesday_percentage
            elif day == "Wednesday" and slot in ["2:30-4:00 PM", "4:00-5:30 PM"]:
                slot_percentages[(day, slot)] = zero_percentage
            elif day == "Friday" and slot == "11:30-1:00 PM":
                slot_percentages[(day, slot)] = zero_percentage
            else:
                slot_percentages[(day, slot)] = default_percentage
    
    return slot_percentages

def define_course_blocks():
    """
    Define course blocks that should not be scheduled at the same time.
    
    Returns:
        tuple: Lists of course blocks and special blocks.
    """
    course_blocks = [
        ['CS114', 'IS210', 'CS450', 'CS337'],
        ['CS241', 'CS280', 'IS350'],
        ['CS288', 'CS332', 'CS301', 'CS356'],  # Special block (<= 2)
        ['CS341', 'CS350', 'CS351', 'CS331', 'CS375'],  # Special block (<= 2)
        ['CS435', 'CS490', 'CS485', 'CS370', 'CS375'],
        ['CS485', 'CS491', 'CS450', 'CS482'],
        ['CS610', 'CS630', 'CS631', 'CS656', 'DS675', 'CS675', 'CS670'],  # Block-grad-core
        ['DS677', 'DS669', 'DS650', 'CS670', 'CS610', 'CS665', 'CS667', 'CS732', 'DS680'],  # Block-grad-DS+Alg
        ['CS608', 'CS645', 'CS646', 'CS647', 'CS648', 'CS678', 'CS696'],  # Block-grad-cyber
        ['IS455', 'IS645'],
        ['IT220', 'IT230', 'IT240', 'IT302'],
        ['IT256', 'IT266', 'IT286', 'IT360', 'IT380', 'IT383', 'IT386'],
        ['IT120', 'IT240']
    ]
    
    # List of special blocks with <= 2 constraints
    special_blocks = [
        ['CS288', 'CS332', 'CS301', 'CS356'],
        ['CS341', 'CS350', 'CS351', 'CS331', 'CS375']
    ]
    
    return course_blocks, special_blocks

def define_valid_slots_for_course_patterns():
    """
    Define valid start times for different course patterns.
    
    Returns:
        tuple: Valid start times for regular days and Fridays.
    """
    # Define valid start times for graduate-style patterns
    valid_start_times = ["8:30-10:00 AM", "6:00-7:30 PM"]

    # Define additional valid start times for Fridays
    friday_start_times = [
        "8:30-10:00 AM", "1:00-2:30 PM", "2:30-4:00 PM",
        "4:00-5:30 PM", "6:00-7:30 PM", "7:30-9:00 PM"
    ]
    
    return valid_start_times, friday_start_times

def initialize_model():
    """
    Initialize the Gurobi optimization model.
    
    Returns:
        tuple: Model and variables dictionary.
    """
    model = Model("Scheduling")
    variables = {}
    return model, variables

def create_decision_variables(model, variables, df, days, time_slots):
    """
    Create binary decision variables for the scheduling model.
    
    Args:
        model (Model): Gurobi model.
        variables (dict): Dictionary to store variables.
        df (DataFrame): DataFrame with course and instructor data.
        days (list): List of days.
        time_slots (list): List of time slots.
        
    Returns:
        dict: Updated variables dictionary.
    """
    print("Creating decision variables...")
    
    for _, row in df.iterrows():
        course = row['Course']
        instructor = row['Instructor']
        num_sections = int(row['# Sections'])

        if num_sections == 0:
            continue

        # Adjust for "CS435" to have 3 parts instead of 2
        if course == "CS435":
            parts = [1, 2, 3]
        else:
            parts = [1, 2]

        for section_id in range(1, num_sections + 1):
            for part in parts:
                for day in days:
                    for slot in time_slots:
                        var_name = f"X_{course}_{instructor}_{section_id}_{part}_{day}_{slot}"
                        variables[var_name] = model.addVar(vtype=GRB.BINARY, name=var_name)
    
    return variables

def add_unique_assignment_constraints(model, variables, df, days, time_slots):
    """
    Add constraints to ensure each course section part is assigned exactly once.
    
    Args:
        model (Model): Gurobi model.
        variables (dict): Dictionary of decision variables.
        df (DataFrame): DataFrame with course and instructor data.
        days (list): List of days.
        time_slots (list): List of time slots.
    """
    print("Adding unique assignment constraints...")
    
    for _, row in df.iterrows():
        course = row['Course']
        instructor = row['Instructor']
        num_sections = int(row['# Sections'])

        if num_sections == 0:
            continue

        # Adjust for "CS435" to have 3 parts instead of 2
        if course == "CS435":
            parts = [1, 2, 3]
        else:
            parts = [1, 2]

        for section_id in range(1, num_sections + 1):
            for part in parts:
                section_vars = []
                for day in days:
                    for slot in time_slots:
                        var_name = f"X_{course}_{instructor}_{section_id}_{part}_{day}_{slot}"
                        if var_name in variables:
                            var = variables[var_name]
                            section_vars.append(var)

                if section_vars:
                    constraint_name = f"unique_slot_{course}_{instructor}_{section_id}_{part}"
                    model.addConstr(sum(section_vars) == 1, name=constraint_name)

def add_instructor_availability_constraints(model, variables, df, days, time_slots):
    """
    Add constraints to ensure an instructor is not scheduled for more than one section 
    at the same time slot.
    
    Args:
        model (Model): Gurobi model.
        variables (dict): Dictionary of decision variables.
        df (DataFrame): DataFrame with course and instructor data.
        days (list): List of days.
        time_slots (list): List of time slots.
    """
    print("Adding instructor availability constraints...")
    
    for instructor in df['Instructor'].unique():
        for day in days:
            for slot in time_slots:
                instructor_vars = []
                for _, row in df[df['Instructor'] == instructor].iterrows():
                    course = row['Course']
                    num_sections = int(row['# Sections'])
                    if course == "CS435":
                        parts = [1, 2, 3]
                    else:
                        parts = [1, 2]
                    for section_id in range(1, num_sections + 1):
                        for part in parts:
                            var_name = f"X_{course}_{instructor}_{section_id}_{part}_{day}_{slot}"
                            if var_name in variables:
                                var = variables[var_name]
                                instructor_vars.append(var)

                if instructor_vars:
                    constraint_name = f"one_section_per_slot_{instructor}_{day}_{slot}"
                    model.addConstr(sum(instructor_vars) <= 1, name=constraint_name)

def add_time_slot_balance_constraints(model, variables, df, days, time_slots, slot_percentages):
    """
    Add constraints to balance the distribution of classes across time slots.
    
    Args:
        model (Model): Gurobi model.
        variables (dict): Dictionary of decision variables.
        df (DataFrame): DataFrame with course and instructor data.
        days (list): List of days.
        time_slots (list): List of time slots.
        slot_percentages (dict): Dictionary mapping (day, time_slot) to target percentage.
    """
    print("Adding time slot balance constraints...")
    
    # Calculate total number of section parts
    total_section_parts = 2 * sum(int(row['# Sections']) for _, row in df.iterrows()) + df[df["Course"] == "CS435"]["# Sections"].sum()
    print(f"Total section parts: {total_section_parts}")
    
    balance_sum = 0
    
    for (day, slot), percentage in slot_percentages.items():
        slot_vars = [variables[var_name] for var_name in variables if var_name.split('_')[5] == day and var_name.split('_')[6] == slot]
        
        # Adjust max_section_parts_slot by adding the number of pre-scheduled courses
        max_section_parts_slot = math.ceil((percentage / 6) * total_section_parts)
        balance_sum += max_section_parts_slot
        
        if percentage > 0:  # Only add slack variables if the percentage is non-zero
            model.addConstr(
                sum(slot_vars) <= max_section_parts_slot,
                name=f"balance_slot_with_slack_{day}_{slot}"
            )
        else:
            model.addConstr(
                sum(slot_vars) <= max_section_parts_slot,
                name=f"balance_slot_no_slack_{day}_{slot}"
            )
    
    print(f"Sum of the balances is: {balance_sum}")
    
    # Evening slot constraints
    evening_percentage = 0.20  # Evening slots defined as 6:00-7:30 PM and 7:30-9:00 PM
    evening_slots = ["6:00-7:30 PM", "7:30-9:00 PM"]
    
    for slot in evening_slots:
        evening_vars = [variables[var_name] for var_name in variables if var_name.split('_')[6] == slot]
        max_section_parts_evening = (evening_percentage / 6) * total_section_parts
        model.addConstr(sum(evening_vars) <= max_section_parts_evening, name=f"balance_evening_{slot}")

def add_restricted_time_slots_constraints(model, variables, df, days):
    """
    Add constraints to limit the number of restricted time slots an instructor can be assigned.
    
    Args:
        model (Model): Gurobi model.
        variables (dict): Dictionary of decision variables.
        df (DataFrame): DataFrame with course and instructor data.
        days (list): List of days.
    """
    print("Adding restricted time slots constraints...")
    
    # Define the specific time slots for the constraint
    restricted_time_slots = ["8:30-10:00 AM", "10:00-11:30 AM", "6:00-7:30 PM", "7:30-9:00 PM"]
    
    # Add constraint for each instructor on each day
    for instructor in df['Instructor'].unique():
        for day in days:
            # Collect the binary variables corresponding to the restricted time slots
            restricted_vars = []
            for course in df['Course'].unique():
                instructor_courses = df[(df['Instructor'] == instructor) & (df['Course'] == course)]
                
                for _, course_row in instructor_courses.iterrows():
                    num_sections = int(course_row['# Sections'])
                    for section_id in range(1, num_sections + 1):
                        parts = [1, 2, 3] if course == "CS435" else [1, 2]
                        for part in parts:
                            for slot in restricted_time_slots:
                                var_name = f"X_{course}_{instructor}_{section_id}_{part}_{day}_{slot}"
                                if var_name in variables:
                                    restricted_vars.append(variables[var_name])
            
            # Add constraint to ensure that at most 3 of the 4 restricted slots can be assigned
            if restricted_vars:
                model.addConstr(
                    sum(restricted_vars) <= 3,
                    name=f"restricted_time_slots_{instructor}_{day}"
                )

def add_course_pattern_constraints(model, variables, df, days, time_slots, 
                                  valid_start_times, friday_start_times):
    """
    Add constraints for course scheduling patterns (graduate vs undergraduate format).
    
    Args:
        model (Model): Gurobi model.
        variables (dict): Dictionary of decision variables.
        df (DataFrame): DataFrame with course and instructor data.
        days (list): List of days.
        time_slots (list): List of time slots.
        valid_start_times (list): Valid start times for graduate pattern.
        friday_start_times (list): Valid start times for graduate pattern on Fridays.
    """
    print("Adding course pattern constraints...")
    
    # Dictionary to store pattern variables (for quicker lookup)
    y_var_dict = {}
    
    for course in df['Course'].unique():
        for instructor in df['Instructor'].unique():
            course_instructor_rows = df[(df['Course'] == course) & (df['Instructor'] == instructor)]
            if course_instructor_rows.empty:
                continue
                
            for section_id in range(1, course_instructor_rows.iloc[0]['# Sections'] + 1):
                # Create variables for graduate and undergraduate patterns
                grad_var = model.addVar(vtype=GRB.BINARY, name=f"Grad_{course}_{instructor}_{section_id}")
                undergrad_var = model.addVar(vtype=GRB.BINARY, name=f"Undergrad_{course}_{instructor}_{section_id}")
                
                compatible_pairs = []  # Track compatible (day1, slot1) and (day2, slot2) pairs
                
                for day1 in days:
                    for slot1 in time_slots:
                        for day2 in days:
                            for slot2 in time_slots:
                                # Define the binary variable for this (day1, slot1), (day2, slot2) pair
                                y_var_name = f"Y_{course}_{instructor}_{section_id}_{day1}_{slot1}_{day2}_{slot2}"
                                y_var = model.addVar(vtype=GRB.BINARY, name=y_var_name)
                                compatible_pairs.append(y_var)
                                y_var_dict[(day1, slot1, day2, slot2)] = y_var
                                
                                # Graduate pattern (consecutive slots on the same day)
                                if day1 == day2:
                                    if day1 == "Friday":
                                        # Friday-specific constraint
                                        if slot1 in friday_start_times and time_slots.index(slot2) == time_slots.index(slot1) + 1:
                                            model.addConstr(y_var <= grad_var, 
                                                          name=f"grad_pair_friday_{course}_{instructor}_{section_id}_{day1}_{slot1}_{slot2}")
                                        else:
                                            # Disable non-consecutive or invalid start times for Fridays
                                            model.addConstr(y_var == 0, 
                                                          name=f"invalid_grad_friday_{course}_{instructor}_{section_id}_{day1}_{slot1}_{slot2}")
                                    else:
                                        # Other days: Allow only specific start times
                                        if slot1 in valid_start_times and time_slots.index(slot2) == time_slots.index(slot1) + 1:
                                            model.addConstr(y_var <= grad_var, 
                                                          name=f"grad_pair_non_friday_{course}_{instructor}_{section_id}_{day1}_{slot1}_{slot2}")
                                        else:
                                            # Disable non-consecutive or invalid start times for other days
                                            model.addConstr(y_var == 0, 
                                                          name=f"invalid_grad_non_friday_{course}_{instructor}_{section_id}_{day1}_{slot1}_{slot2}")
                                
                                # Undergraduate pattern (same slot, different valid day pairs)
                                elif day1 != day2 and slot1 == slot2:
                                    if (day1 == "Monday" and day2 in ["Wednesday", "Thursday"]) or \
                                       (day1 == "Tuesday" and day2 in ["Thursday", "Friday"]) or \
                                       (day1 == "Wednesday" and day2 == "Friday"):
                                        model.addConstr(y_var <= undergrad_var, 
                                                      name=f"undergrad_pair_enforce_{course}_{instructor}_{section_id}_{day1}_{slot1}_{day2}_{slot2}")
                                    else:
                                        model.addConstr(y_var == 0, 
                                                      name=f"invalid_undergrad_pair_{course}_{instructor}_{section_id}_{day1}_{slot1}_{day2}_{slot2}")
                                else:
                                    model.addConstr(y_var == 0, 
                                                  name=f"invalid_pair_{course}_{instructor}_{section_id}_{day1}_{slot1}_{day2}_{slot2}")
                
                # Ensure Part 1 of courses where the third character is "7" cannot be scheduled from 8:30 to 10:00 AM
                for day1 in days:
                    for slot1 in time_slots:
                        if course[2] == "7" and slot1 == "8:30-10:00 AM":
                            var_part1 = f"X_{course}_{instructor}_{section_id}_1_{day1}_{slot1}"
                            if var_part1 in variables:
                                model.addConstr(variables[var_part1] == 0, 
                                              name=f"no_8_30_to_10_CS7XX_{course}_{instructor}_{section_id}_{day1}_{slot1}")
                
                # Link Part 1 variables to the sum of corresponding Y variables
                for day1 in days:
                    for slot1 in time_slots:
                        var_part1 = f"X_{course}_{instructor}_{section_id}_1_{day1}_{slot1}"
                        if var_part1 in variables:
                            y_vars_for_part1 = [y_var_dict.get((day1, slot1, day2, slot2), None) 
                                              for day2 in days for slot2 in time_slots 
                                              if (day1, slot1, day2, slot2) in y_var_dict]
                            if y_vars_for_part1:
                                model.addConstr(sum(y_vars_for_part1) == variables[var_part1], 
                                              name=f"part1_link_{course}_{instructor}_{section_id}_{day1}_{slot1}")
                
                # Link Part 2 variables to the sum of corresponding Y variables
                for day2 in days:
                    for slot2 in time_slots:
                        var_part2 = f"X_{course}_{instructor}_{section_id}_2_{day2}_{slot2}"
                        if var_part2 in variables:
                            y_vars_for_part2 = [y_var_dict.get((day1, slot1, day2, slot2), None) 
                                              for day1 in days for slot1 in time_slots 
                                              if (day1, slot1, day2, slot2) in y_var_dict]
                            if y_vars_for_part2:
                                model.addConstr(sum(y_vars_for_part2) == variables[var_part2], 
                                              name=f"part2_link_{course}_{instructor}_{section_id}_{day2}_{slot2}")
                
                # Ensure exactly one (day1, slot1), (day2, slot2) pair is selected
                model.addConstr(sum(compatible_pairs) == 1, name=f"select_one_pair_{course}_{instructor}_{section_id}")
                
                # Ensure that only one pattern (graduate or undergraduate) is chosen
                model.addConstr(grad_var + undergrad_var == 1, name=f"select_one_pattern_{course}_{instructor}_{section_id}")

def add_course_block_constraints(model, variables, df, days, time_slots, course_blocks, special_blocks):
    """
    Add constraints to prevent courses in the same block from being scheduled at the same time.
    
    Args:
        model (Model): Gurobi model.
        variables (dict): Dictionary of decision variables.
        df (DataFrame): DataFrame with course and instructor data.
        days (list): List of days.
        time_slots (list): List of time slots.
        course_blocks (list): List of course blocks.
        special_blocks (list): List of special blocks that allow up to 2 courses at the same time.
    """
    print("Adding course block constraints...")
    
    for block in course_blocks:
        for course1 in block:
            for course2 in block:
                if course1 != course2:
                    for instructor1 in df[df['Course'] == course1]['Instructor'].unique():
                        for instructor2 in df[df['Course'] == course2]['Instructor'].unique():
                            for day in days:
                                for slot in time_slots:
                                    # Check if the current block is a special block (<= 2 constraints)
                                    max_constraint = 2 if block in special_blocks else 1
                                    
                                    for part in [1, 2]:  # We only need to check parts 1 and 2 for all courses
                                        # Create a list of sections for each course-instructor combination
                                        sections1 = df[(df['Course'] == course1) & (df['Instructor'] == instructor1)]['# Sections'].sum()
                                        sections2 = df[(df['Course'] == course2) & (df['Instructor'] == instructor2)]['# Sections'].sum()
                                        
                                        # Collect all variables for these course-instructor-part combinations
                                        vars1 = []
                                        vars2 = []
                                        
                                        for section_id1 in range(1, int(sections1) + 1):
                                            var_name1 = f"X_{course1}_{instructor1}_{section_id1}_{part}_{day}_{slot}"
                                            if var_name1 in variables:
                                                vars1.append(variables[var_name1])
                                                
                                        for section_id2 in range(1, int(sections2) + 1):
                                            var_name2 = f"X_{course2}_{instructor2}_{section_id2}_{part}_{day}_{slot}"
                                            if var_name2 in variables:
                                                vars2.append(variables[var_name2])
                                        
                                        # If both courses have variables for this part, day, and slot
                                        if vars1 and vars2:
                                            model.addConstr(
                                                sum(vars1) + sum(vars2) <= max_constraint,
                                                name=f"block_constraint_{course1}_{course2}_{day}_{slot}_part{part}"
                                            )

def add_evening_constraints(model, variables, df, days):
    """
    Add constraints for evening class scheduling.
    
    Args:
        model (Model): Gurobi model.
        variables (dict): Dictionary of decision variables.
        df (DataFrame): DataFrame with course and instructor data.
        days (list): List of days.
    """
    print("Adding evening scheduling constraints...")
    
    part1_slot = "6:00-7:30 PM"
    part2_slot = "7:30-9:00 PM"
    
    # Iterate over the DataFrame and add constraints for each section
    for _, row in df.iterrows():
        course = row['Course']
        instructor = row['Instructor']
        num_sections = int(row['# Sections'])
        
        if num_sections == 0:
            continue

        for section_id in range(1, num_sections + 1):
            for day in days:
                # Get the variable names for part 1 and part 2 for the relevant time slots
                part1_var_name = f"X_{course}_{instructor}_{section_id}_1_{day}_{part1_slot}"
                part2_var_name = f"X_{course}_{instructor}_{section_id}_2_{day}_{part2_slot}"
                
                # Ensure both variables exist
                if part1_var_name in variables and part2_var_name in variables:
                    # Ensure that the assignment of part 1 at 6:00-7:30 PM equals the assignment of part 2 at 7:30-9:00 PM
                    model.addConstr(variables[part1_var_name] == variables[part2_var_name],
                        name=f"timing_constraint_{course}_{instructor}_{section_id}_{day}")

def add_consecutive_slots_constraints(model, variables, df, days, time_slots):
    """
    Add constraints to prevent instructors from teaching more than 2 consecutive time slots.
    
    Args:
        model (Model): Gurobi model.
        variables (dict): Dictionary of decision variables.
        df (DataFrame): DataFrame with course and instructor data.
        days (list): List of days.
        time_slots (list): List of time slots.
    """
    print("Adding consecutive slots constraints...")
    
    for instructor in df['Instructor'].unique():
        # Iterate over all days
        for day in days:
            # Iterate over all possible starting time slots (the first of three consecutive slots)
            for i in range(len(time_slots) - 2):
                slot1 = time_slots[i]
                slot2 = time_slots[i + 1]
                slot3 = time_slots[i + 2]
                
                # Filter the DataFrame to get only the courses and sections taught by the current instructor
                instructor_df = df[df['Instructor'] == instructor]

                # Create variables for each possible combination of courses and sections in the three slots
                for idx1, row1 in instructor_df.iterrows():
                    course1 = row1['Course']
                    num_sections1 = int(row1['# Sections'])
                    
                    if num_sections1 == 0:
                        continue

                    # Loop over the first section and part for the first slot
                    for section_id1 in range(1, num_sections1 + 1):
                        parts1 = [1, 2, 3] if course1 == "CS435" else [1, 2]
                        for part1 in parts1:
                            # Build the variable name for the first slot
                            var_name1 = f"X_{course1}_{instructor}_{section_id1}_{part1}_{day}_{slot1}"

                            # Loop over the second course taught by the instructor
                            for idx2, row2 in instructor_df.iterrows():
                                course2 = row2['Course']
                                num_sections2 = int(row2['# Sections'])
                                
                                if num_sections2 == 0:
                                    continue

                                # Loop over the second section and part for the second slot
                                for section_id2 in range(1, num_sections2 + 1):
                                    parts2 = [1, 2, 3] if course2 == "CS435" else [1, 2]
                                    for part2 in parts2:
                                        # Build the variable name for the second slot
                                        var_name2 = f"X_{course2}_{instructor}_{section_id2}_{part2}_{day}_{slot2}"

                                        # Loop over the third course taught by the instructor
                                        for idx3, row3 in instructor_df.iterrows():
                                            course3 = row3['Course']
                                            num_sections3 = int(row3['# Sections'])
                                            
                                            if num_sections3 == 0:
                                                continue

                                            # Loop over the third section and part for the third slot
                                            for section_id3 in range(1, num_sections3 + 1):
                                                parts3 = [1, 2, 3] if course3 == "CS435" else [1, 2]
                                                for part3 in parts3:
                                                    # Build the variable name for the third slot
                                                    var_name3 = f"X_{course3}_{instructor}_{section_id3}_{part3}_{day}_{slot3}"

                                                    # Check if all variables are different 
                                                    # (not the same course/section/part combination)
                                                    if (course1 != course2 or section_id1 != section_id2 or part1 != part2) and \
                                                       (course2 != course3 or section_id2 != section_id3 or part2 != part3) and \
                                                       (course1 != course3 or section_id1 != section_id3 or part1 != part3):
                                                        
                                                        consecutive_sum = 0
                                                        # Add the variables to the sum if they exist
                                                        if var_name1 in variables:
                                                            consecutive_sum += variables[var_name1]
                                                        if var_name2 in variables:
                                                            consecutive_sum += variables[var_name2]
                                                        if var_name3 in variables:
                                                            consecutive_sum += variables[var_name3]

                                                        # Add the constraint that the sum of these variables must be <= 2
                                                        if consecutive_sum > 0:  # Only add if there are variables to constrain
                                                            model.addConstr(
                                                                consecutive_sum <= 2, 
                                                                name=f"consecutive_slots_{instructor}_{day}_{slot1}_{slot2}_{slot3}"
                                                            )

def add_restricted_monday_constraints(model, variables, df, section_capacity_map):
    """
    Add constraints for restricted Monday slots.
    
    Args:
        model (Model): Gurobi model.
        variables (dict): Dictionary of decision variables.
        df (DataFrame): DataFrame with course and instructor data.
        section_capacity_map (dict): Dictionary mapping (course, instructor, section_id) to capacity.
    """
    print("Adding restricted Monday constraints...")
    
    restricted_day = "Monday"
    restricted_time_slot = "4:00-5:30 PM"
    
    # Add the constraint: Only courses with course number > 199 and capacity < 35 can be scheduled on Monday from 4:00 to 5:30
    for _, row in df.iterrows():
        course = row['Course']
        instructor = row['Instructor']
        course_number = int(row['Course_Number'])
        
        for sc in range(1, int(row['# Sections']) + 1):
            capacity = section_capacity_map.get((course, instructor, sc))

            # Check if the course meets the condition for being scheduled in this restricted slot
            if course_number > 199 and capacity < 35:
                # Loop over sections and parts to ensure the variables for this course are allowed to be scheduled
                for section_id in range(1, int(row['# Sections']) + 1):
                    parts = [1, 2, 3] if course == "CS435" else [1, 2]
                    for part in parts:
                        var_name = f"X_{course}_{instructor}_{section_id}_{part}_{restricted_day}_{restricted_time_slot}"
                        if var_name in variables:
                            # No constraint needed, the course meets the condition
                            continue
            else:
                # If the course does not meet the conditions, add a constraint to prevent it from being scheduled
                for section_id in range(1, int(row['# Sections']) + 1):
                    parts = [1, 2, 3] if course == "CS435" else [1, 2]
                    for part in parts:
                        var_name = f"X_{course}_{instructor}_{section_id}_{part}_{restricted_day}_{restricted_time_slot}"
                        if var_name in variables:
                            model.addConstr(
                                variables[var_name] == 0, 
                                name=f"restricted_slot_{course}_{instructor}_{section_id}_{part}_{restricted_day}_{restricted_time_slot}"
                            )

def add_pre_scheduled_constraints(model, variables, df_pre_scheduled):
    """
    Add constraints to enforce pre-scheduled courses.
    
    Args:
        model (Model): Gurobi model.
        variables (dict): Dictionary of decision variables.
        df_pre_scheduled (DataFrame): DataFrame with pre-scheduled course information.
    """
    print("Adding pre-scheduled constraints...")
    
    # Enforce that no other courses for this instructor can be scheduled at the specified (day, time)
    for _, row in df_pre_scheduled.iterrows():
        instructor = row['Instructor']
        day = row['Day']
        times = row['Time']

        # Set all variables for this (instructor, day, time) to 0
        for var_name in variables:
            var_parts = var_name.split('_')
            # Check if it's a valid X variable
            if len(var_parts) >= 7:
                var_instructor = var_parts[2]
                var_day = var_parts[5]
                var_time = var_parts[6]

                # Check if the variable matches the instructor, day, and time
                if var_instructor == instructor and var_day == day and var_time == times:
                    model.addConstr(
                        variables[var_name] == 0, 
                        name=f"block_{instructor}_{day}_{times}"
                    )

def add_health_religion_constraints(model, variables, df, df_constraints, time_slot_mapping, 
                                    time_slot_index, total_points):
    """
    Add constraints for health and religion preferences with high penalties.
    
    Args:
        model (Model): Gurobi model.
        variables (dict): Dictionary of decision variables.
        df (DataFrame): DataFrame with course and instructor data.
        df_constraints (DataFrame): DataFrame with constraints and preferences.
        time_slot_mapping (dict): Dictionary mapping day abbreviations to full day names.
        time_slot_index (dict): Dictionary mapping time slot indices to time slot names.
        total_points (LinExpr): Expression tracking total objective function points.
        
    Returns:
        LinExpr: Updated total_points expression.
    """
    print("Adding health and religion constraints...")
    
    for _, row in df_constraints.iterrows():
        instructor_info = row['Instructor UCID: Type']
        slots = row['Slots']
        if isinstance(instructor_info, float):
            break

        # Parse the instructor UCID and type
        email, constraint_type = instructor_info.split(": ")

        # We only care about the instructors with "Health" or "Religion" type
        if constraint_type.strip() in ["Health", "Religion"]:
            # Parse the blocked time slots for this instructor
            blocked_slots = slots.split("|")[1:-1]  # Remove empty elements from split

            for slot_code in blocked_slots:
                # Extract the day and time slot
                day_abbrev = slot_code[0]  # M, T, W, R, F
                time_slot_num = slot_code[1]  # 1-8

                day_full = time_slot_mapping[day_abbrev]
                time_slot_full = time_slot_index[time_slot_num]

                # Find the instructor name from the df
                instructor_row = df[df['Email'] == email]

                if not instructor_row.empty:
                    instructor_name = instructor_row['Instructor'].iloc[0]  # Get the name as a scalar

                    # Add a soft constraint to block this time slot for all parts of the instructor's courses
                    for course in df[df['Instructor'] == instructor_name]['Course']:
                        filtered_df = df[df['Course'] == course]

                        if not filtered_df.empty:  # Only proceed if the filtered DataFrame is not empty
                            num_sections = int(filtered_df['# Sections'].iloc[0])

                            for section_id in range(1, num_sections + 1):
                                parts = [1, 2, 3] if course == "CS435" else [1, 2]  # CS435 has 3 parts

                                for part in parts:
                                    var_name = f"X_{course}_{instructor_name}_{section_id}_{part}_{day_full}_{time_slot_full}"

                                    # Ensure the variable exists in the model
                                    if var_name in variables:
                                        slack_var_name = f"Slack_{course}_{instructor_name}_{section_id}_{part}_{day_full}_{time_slot_full}"
                                        slack_var = model.addVar(vtype=GRB.BINARY, name=slack_var_name)

                                        # Add the soft constraint (allow slack)
                                        model.addConstr(
                                            variables[var_name] <= slack_var,
                                            name=f"health_religion_constraint_{instructor_name}_{day_full}_{time_slot_full}"
                                        )

                                        # Subtract 2048 points if the constraint is violated
                                        total_points -= 2048 * slack_var
    
    return total_points

def add_instructor_preference_constraints(model, variables, df, df_constraints, time_slot_mapping, 
                                          time_slot_index, total_points):
    """
    Add constraints for instructor preferences with appropriate penalties.
    
    Args:
        model (Model): Gurobi model.
        variables (dict): Dictionary of decision variables.
        df (DataFrame): DataFrame with course and instructor data.
        df_constraints (DataFrame): DataFrame with constraints and preferences.
        time_slot_mapping (dict): Dictionary mapping day abbreviations to full day names.
        time_slot_index (dict): Dictionary mapping time slot indices to time slot names.
        total_points (LinExpr): Expression tracking total objective function points.
        
    Returns:
        LinExpr: Updated total_points expression.
    """
    print("Adding instructor preference constraints...")
    
    instructor_penalty_tracker = {}
    instructor_soft_violated = []

    for idx, row in df_constraints.iterrows():
        instructor_info = row['Instructor UCID: Type']
        slots = row['Slots']
        if isinstance(instructor_info, float):
            break
        
        # Parse the instructor UCID and constraint type
        email, constraint_type = instructor_info.split(": ")
        
        # Parse the blocked time slots for this instructor
        blocked_slots = slots.split("|")[1:-1]  # Remove empty elements from split

        # Handle different types of constraints
        if constraint_type.strip() == "Pref-1":
            points = 8
        elif constraint_type.strip() == "Pref-2":
            points = 4
        elif constraint_type.strip() == "Pref-3":
            points = 2
        elif constraint_type.strip() in ["Health", "Religion"]:
            continue  # No points assigned for Health/Religion as these are hard constraints
        elif constraint_type.strip() == "Childcare":
            points = -1024
        else:
            points = -8  # Default negative points for all other types

        for slot_code in blocked_slots:
            try:
                # Check if the length of slot_code is at least 2
                if len(slot_code) < 2:
                    print(f"Error in row {idx}, email: {email}, constraint: {constraint_type}, slot code: '{slot_code}' (invalid length)")
                    continue  # Skip this slot if it's too short

                # Extract the day and time slot
                day_abbrev = slot_code[0]  # M, T, W, R, F
                time_slot_num = slot_code[1]  # 1-8

                # Extract the full day and time slot
                day_full = time_slot_mapping[day_abbrev]
                time_slot_full = time_slot_index[time_slot_num]

                # Find the instructor's name using the email, skip if not found
                instructor_row = df[df['Email'] == email]
                if instructor_row.empty:
                    continue

                instructor_name = instructor_row['Instructor'].iloc[0]

                # Add or subtract points based on the constraint type for all sections and parts of the instructor's courses
                for course in df[df['Instructor'] == instructor_name]['Course']:
                    filtered_df = df[df['Course'] == course]
                    if not filtered_df.empty:  # Only proceed if the filtered DataFrame is not empty
                        num_sections = int(filtered_df[
                            (filtered_df['Instructor'] == instructor_name) &
                            (filtered_df['Course'] == course)
                        ]['# Sections'].iloc[0])

                        for section_id in range(1, num_sections + 1):
                            parts = [1, 2, 3] if course == "CS435" else [1, 2]
                            for part in parts:  # 2 parts per section (3 for CS435)
                                # Variable name now excludes constraint type
                                var_name = f"X_{course}_{instructor_name}_{section_id}_{part}_{day_full}_{time_slot_full}"
                                # Ensure the variable exists in the model
                                if var_name in variables:
                                    total_points += points * variables[var_name]
                                    if points < -1:
                                        instructor_soft_violated.append((var_name, points))
            except IndexError:
                # Print the row index and the problematic slot
                print(f"Error in row {idx}, email: {email}, constraint: {constraint_type}, slot code: '{slot_code}'")
                continue  # Skip this slot and move to the next one
    
    return total_points, instructor_soft_violated

def add_teaching_days_variables(model, variables, df, days, time_slots):
    """
    Add variables to track which days each instructor teaches.
    
    Args:
        model (Model): Gurobi model.
        variables (dict): Dictionary of decision variables.
        df (DataFrame): DataFrame with course and instructor data.
        days (list): List of days.
        time_slots (list): List of time slots.
        
    Returns:
        dict: Dictionary of Z variables tracking teaching days.
    """
    print("Adding teaching days tracking variables...")
    
    z_vars = {}
    for instructor in df['Instructor'].unique():
        for day in days:
            # Binary variable indicating whether the instructor teaches on that day
            z_var = model.addVar(vtype=GRB.BINARY, name=f"Z_{instructor}_{day}")
            z_vars[(instructor, day)] = z_var

            # Relevant X variables for this instructor and day
            relevant_x_vars = []
            for course in df['Course'].unique():
                course_instructor_rows = df[(df['Course'] == course) & (df['Instructor'] == instructor)]

                if course_instructor_rows.empty:
                    continue

                num_sections = course_instructor_rows['# Sections'].iloc[0]
                for section_id in range(1, int(num_sections) + 1):
                    parts = [1, 2, 3] if course == "CS435" else [1, 2]
                    for part in parts:
                        for slot in time_slots:
                            x_var_name = f"X_{course}_{instructor}_{section_id}_{part}_{day}_{slot}"
                            if x_var_name in variables:
                                x_var = variables[x_var_name]
                                relevant_x_vars.append(x_var)
                                # Ensure that if X variable is 1, z_var must be 1
                                model.addConstr(
                                    x_var <= z_var,
                                    name=f"x_var_link_{course}_{instructor}_{section_id}_{part}_{day}_{slot}"
                                )

            # If no relevant X variables are found, ensure that z_var is set to 0
            if not relevant_x_vars:
                model.addConstr(z_var == 0, name=f"no_classes_{instructor}_{day}")
    
    return z_vars

def add_day_preference_penalties(model, z_vars, general_preferences_df, df):
    """
    Add penalties based on day preferences.
    
    Args:
        model (Model): Gurobi model.
        z_vars (dict): Dictionary of Z variables tracking teaching days.
        general_preferences_df (DataFrame): DataFrame with general preferences.
        df (DataFrame): DataFrame with course and instructor data.
        
    Returns:
        LinExpr: Expression for day preference penalties.
    """
    print("Adding day preference penalties...")
    
    # Calculate penalty sum S
    day_penalty_sum = LinExpr()
    
    # Create a dictionary mapping emails to their day preferences
    day_preference_dict = general_preferences_df.set_index('Email')['Day Preference'].to_dict()

    # Iterate over instructors and days to add penalties based on their preferences
    for instructor in df['Instructor'].unique():
        # Get the email of the instructor
        email = df[df['Instructor'] == instructor]['Email'].iloc[0]

        # Check if the instructor prefers condensed days
        prefers_condensed_days = day_preference_dict.get(email, "No") == "I prefer to condense my sections into fewer days"

        # Set the penalty value
        penalty_value = -8 if prefers_condensed_days else -3

        for day in ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']:
            # Binary variable indicating whether the instructor teaches on that day
            z_var = z_vars.get((instructor, day))  # This should now always exist
            if z_var:
                # Add penalty to the penalty sum
                day_penalty_sum += penalty_value * z_var
    
    return day_penalty_sum

def add_consecutive_preference_penalties(model, variables, df, consecutive_preference, days, time_slots):
    """
    Add penalties based on preferences for consecutive slots.
    
    Args:
        model (Model): Gurobi model.
        variables (dict): Dictionary of decision variables.
        df (DataFrame): DataFrame with course and instructor data.
        consecutive_preference (dict): Dictionary mapping emails to consecutive slot preferences.
        days (list): List of days.
        time_slots (list): List of time slots.
        
    Returns:
        LinExpr: Expression for consecutive slot preference penalties.
    """
    print("Adding consecutive slot preference penalties...")
    
    # Initialize a penalty expression to sum penalties
    consecutive_penalty_sum = LinExpr()
    
    # Default penalty value
    penalty_value = -2048  # High penalty for consecutive slots if disliked

    for email, prefers_consecutive in consecutive_preference.items():
        if prefers_consecutive == "No":  # If they dislike consecutive slots
            # Get the instructor's name using the email from the DataFrame
            instructor_row = df[df['Email'] == email]
            if instructor_row.empty:
                continue
            instructor_name = instructor_row['Instructor'].iloc[0]

            for day in days:
                for slot_idx in range(len(time_slots) - 1):
                    slot1 = time_slots[slot_idx]
                    slot2 = time_slots[slot_idx + 1]
                    
                    for course in df['Course'].unique():
                        relevant_rows = df[(df['Instructor'] == instructor_name) & (df['Course'] == course)]
                        if relevant_rows.empty:
                            continue
                        
                        num_sections = int(relevant_rows['# Sections'].iloc[0])
                        for section_id in range(1, num_sections + 1):
                            # Define X variables for consecutive slots
                            x_var1_name = f"X_{course}_{instructor_name}_{section_id}_1_{day}_{slot1}"
                            x_var2_name = f"X_{course}_{instructor_name}_{section_id}_2_{day}_{slot2}"
                            
                            if x_var1_name in variables and x_var2_name in variables:
                                x_var1 = variables[x_var1_name]
                                x_var2 = variables[x_var2_name]

                                # Add penalty variable
                                penalty_var = model.addVar(
                                    vtype=GRB.BINARY, 
                                    name=f"Penalty_{instructor_name}_{day}_{slot1}_{slot2}"
                                )
                                model.addConstr(
                                    x_var1 + x_var2 - 2 * penalty_var <= 1, 
                                    name=f"consecutive_penalty_{instructor_name}_{day}_{slot1}_{slot2}"
                                )
                                
                                # Add penalty to the total penalty sum
                                consecutive_penalty_sum += penalty_value * penalty_var
    
    return consecutive_penalty_sum

def add_format_preference_penalties(model, variables, df, general_preferences_df, days, time_slots):
    """
    Add penalties based on course format preferences.
    
    Args:
        model (Model): Gurobi model.
        variables (dict): Dictionary of decision variables.
        df (DataFrame): DataFrame with course and instructor data.
        general_preferences_df (DataFrame): DataFrame with general preferences.
        days (list): List of days.
        time_slots (list): List of time slots.
        
    Returns:
        LinExpr: Expression for format preference penalties.
    """
    print("Adding format preference penalties...")
    
    # Initialize a penalty expression for violated preferences
    format_penalty_sum = LinExpr()
    
    # Penalty for violating the format preference
    penalty_value = -8
    
    # Iterate over instructor preferences and add penalties for violations
    for _, row in general_preferences_df.iterrows():
        email = row['Email']
        preference = row['Preference']

        # Find the instructor's name from the main DataFrame using the email
        instructor_row = df[df['Email'] == email]
        if instructor_row.empty:
            continue

        instructor_name = instructor_row['Instructor'].iloc[0]

        # Filter DataFrame for courses assigned to the instructor
        instructor_courses = df[df['Instructor'] == instructor_name]

        for _, course_row in instructor_courses.iterrows():
            course = course_row['Course']
            num_sections = int(course_row['# Sections'])

            for section_id in range(1, num_sections + 1):
                parts = [1, 2, 3] if course == "CS435" else [1, 2]

                # Iterate over pairs of time slots
                for day in days:
                    for i in range(len(time_slots) - 1):  # Ensure next slot exists
                        slot1 = time_slots[i]
                        slot2 = time_slots[i + 1]

                        part1_var_name = f"X_{course}_{instructor_name}_{section_id}_1_{day}_{slot1}"
                        part2_var_name = f"X_{course}_{instructor_name}_{section_id}_2_{day}_{slot2}"

                        if part1_var_name in variables and part2_var_name in variables:
                            part1_var = variables[part1_var_name]
                            part2_var = variables[part2_var_name]

                            # Add penalty variables for violations
                            penalty_var = model.addVar(
                                vtype=GRB.BINARY, 
                                name=f"Penalty_{instructor_name}_{day}_{slot1}_{slot2}"
                            )

                            if preference == "3-hour format":
                                # Add penalty for non-consecutive parts (violation of 3-hour format preference)
                                model.addConstr(
                                    part1_var - part2_var <= penalty_var, 
                                    name=f"violation_consecutive_{course}_{instructor_name}_{section_id}_{day}_{slot1}_{slot2}"
                                )
                                model.addConstr(
                                    part2_var - part1_var <= penalty_var, 
                                    name=f"violation_consecutive_{course}_{instructor_name}_{section_id}_{day}_{slot1}_{slot2}"
                                )

                            elif preference == "1.5+1.5 hour format":
                                # Add penalty for assigning consecutive parts (violation of non-consecutive preference)
                                model.addConstr(
                                    part1_var + part2_var - penalty_var <= 1, 
                                    name=f"violation_non_consecutive_{course}_{instructor_name}_{section_id}_{day}_{slot1}_{slot2}"
                                )

                            # Add the penalty to the sum
                            format_penalty_sum += penalty_value * penalty_var
    
    return format_penalty_sum

def build_objective_function(model, total_points, format_penalty_sum, day_penalty_sum, consecutive_penalty_sum):
    """
    Build the objective function for the model.
    
    Args:
        model (Model): Gurobi model.
        total_points (LinExpr): Expression for total points.
        format_penalty_sum (LinExpr): Expression for format preference penalties.
        day_penalty_sum (LinExpr): Expression for day preference penalties.
        consecutive_penalty_sum (LinExpr): Expression for consecutive slot preference penalties.
    """
    print("Building objective function...")
    
    # Combine all components of the objective function
    objective = total_points + format_penalty_sum + day_penalty_sum + consecutive_penalty_sum
    
    # Set the model's objective to maximize total points
    model.setObjective(objective, GRB.MAXIMIZE)
    
    # Update the model to incorporate the objective
    model.update()

def solve_model(model, time_limit=600):
    """
    Solve the optimization model.
    
    Args:
        model (Model): Gurobi model.
        time_limit (int): Time limit for optimization in seconds.
        
    Returns:
        int: Status code of the optimization.
    """
    print(f"Optimizing the model with time limit of {time_limit} seconds...")
    
    # Set time limit
    model.setParam("TimeLimit", time_limit)
    
    # Solve the model
    start_optimize = time.time()
    model.optimize()
    optimization_time = time.time() - start_optimize
    
    print(f"Optimization completed in {optimization_time:.2f} seconds")
    
    return model.Status

def analyze_solution(model, variables, df, days, time_slots, section_capacity_map, instructor_soft_violated):
    """
    Analyze the solution and prepare output.
    
    Args:
        model (Model): Gurobi model.
        variables (dict): Dictionary of decision variables.
        df (DataFrame): DataFrame with course and instructor data.
        days (list): List of days.
        time_slots (list): List of time slots.
        section_capacity_map (dict): Dictionary mapping (course, instructor, section) to capacity.
        instructor_soft_violated (list): List of instructor soft constraints that were violated.
        
    Returns:
        tuple: Schedules sorted by course and by instructor.
    """
    print("Analyzing solution...")
    
    # If model is infeasible, return early
    if model.Status == GRB.INFEASIBLE:
        print("The model is infeasible. Check constraints for conflicts.")
        return None, None
    
    # Extract schedule from solution
    schedule = []
    slack_values = []
    
    # Dictionary to track section numbers for each course
    course_section_tracker = {}

    for var in model.getVars():
        if var.varName.startswith('Slack_') and var.x > 0.5:
            # Store violated slack variables
            slack_values.append((var.varName, var.x))
        elif var.varName.startswith('X_') and var.x > 0.5:
            # Extract course, instructor, section, part, day, slot from variable name
            _, course, instructor, section_id, part, day, slot = var.varName.split('_')
            
            # Find the corresponding email for the instructor
            email = df.loc[df['Instructor'] == instructor, 'Email'].values[0]
            
            schedule.append((course, instructor, email, section_id, part, day, slot))
    
    # Sort schedule by course, instructor, section, part, day, slot
    schedule.sort()
    
    # Print violated slack variables (usually health/religion constraints)
    print("\nViolated Slack Variables (costing -2048 points):\n")
    for var_name, value in slack_values:
        print(f"{var_name}: Value {value}")
    
    print(f"\nTotal number of violated slack variables: {len(slack_values)}")
    
    # Analyze time slot distribution
    calculate_scheduled_percentages(df, variables, model, days, time_slots)
    
    # Create instructor day tracking
    instructor_days = defaultdict(set)
    for entry in schedule:
        course, instructor, email, section_id, part, day, slot = entry
        instructor_days[instructor].add(day)
    
    # Create a list with instructor day counts
    instructor_day_counts = []
    for instructor, days_set in instructor_days.items():
        num_days = len(days_set)
        instructor_day_counts.append((instructor, num_days, sorted(days_set)))
    
    # Sort by number of days (descending)
    instructor_day_counts_sorted = sorted(instructor_day_counts, key=lambda x: x[1], reverse=True)
    
    return schedule, instructor_day_counts_sorted

def generate_output_files(schedule, instructor_day_counts, section_capacity_map, timestamp=None):
    """
    Generate output files with the schedule information.
    
    Args:
        schedule (list): List of scheduled courses.
        instructor_day_counts (list): List of instructors with day counts.
        section_capacity_map (dict): Dictionary mapping (course, instructor, section) to capacity.
        timestamp (str): Optional timestamp for file naming.
    """
    print("Generating output files...")
    
    if timestamp is None:
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    
    # Dictionary to track section numbers for each course
    course_section_tracker = {}
    
    # Write schedule sorted by course
    with open(f"final_schedule_sorted_by_course_{timestamp}.txt", "w") as f:
        f.write("Course Schedule (Lexicographically Sorted):\n\n")
        for entry in schedule:
            course, instructor, email, section_id, part, day, slot = entry
            course = course.strip()
            instructor = instructor.strip()

            # Track the section number for each course
            if course not in course_section_tracker:
                course_section_tracker[course] = 1  # Initialize section number for this course
            
            # Assign the next available section number for this course
            assigned_section_number = course_section_tracker[course]
            
            # Retrieve the capacity from the map
            capacity = section_capacity_map.get((course, instructor, int(section_id)))

            # Write the information to the file
            f.write(f"Course: {course}, Instructor: {instructor}, Email: {email}, Section: {assigned_section_number}, "
                    f"Part: {part}, Day: {day}, Slot: {slot}, Capacity: {capacity}\n")
            
            # Increment the section tracker after part 2 of a regular course or part 3 of CS435
            if (part == "2" and course != "CS435") or (part == "3" and course == "CS435"):
                course_section_tracker[course] += 1
                f.write(f"\n")
    
    print(f"Final schedule with slack values written to final_schedule_sorted_by_course_{timestamp}.txt")

    # Sort the schedule based on the instructor's name
    schedule_sorted_by_instructor = sorted(schedule, key=lambda x: x[1])

    # Write the instructor-sorted schedule
    with open(f"final_schedule_sorted_by_instructor_{timestamp}.txt", "w") as f:
        f.write("Course Schedule (Sorted by Instructor):\n\n")
        course_section_tracker = {}
        current_instructor = None
        
        for entry in schedule_sorted_by_instructor:
            course, instructor, email, section_id, part, day, slot = entry
            
            # Track the section number for each course
            if course not in course_section_tracker:
                course_section_tracker[course] = 1
                
            assigned_section_number = course_section_tracker[course]

            # If we encounter a new instructor, print their email and name first
            if instructor != current_instructor:
                if current_instructor is not None:
                    f.write("\n")  # Separate different instructors' sections
                
                f.write(f"Instructor: {instructor}, Email: {email}\n")
                current_instructor = instructor

            capacity = section_capacity_map.get((course, instructor, int(section_id)))

            # Write the course details for the current instructor
            f.write(f"\tCourse: {course}, Section: {assigned_section_number}, Part: {part}, "
                    f"Day: {day}, Slot: {slot}, Capacity: {capacity}\n")
            
            # Increment the section tracker after part 2 of a regular course or part 3 of CS435
            if (part == "2" and course != "CS435") or (part == "3" and course == "CS435"):
                course_section_tracker[course] += 1
    
    print(f"Final schedule sorted by instructor written to final_schedule_sorted_by_instructor_{timestamp}.txt")
    
    # Write instructor day counts
    with open(f"instructors_sorted_by_days_on_campus_{timestamp}.txt", "w") as f:
        f.write("Instructors sorted by the number of days they come to campus:\n\n")
        
        for instructor, num_days, days_list in instructor_day_counts:
            # Write the instructor, number of days, and the days they have a class
            f.write(f"Instructor: {instructor}, Number of Days: {num_days}, Days: {', '.join(days_list)}\n")
    
    print(f"Instructors sorted by number of days written to instructors_sorted_by_days_on_campus_{timestamp}.txt")

def calculate_scheduled_percentages(df, variables, model, days, time_slots):
    """
    Calculate the percentage of scheduled classes on each day and time slot.
    
    Args:
        df (DataFrame): DataFrame with course and instructor data.
        variables (dict): Dictionary of decision variables.
        model (Model): Gurobi model.
        days (list): List of days.
        time_slots (list): List of time slots.
        
    Returns:
        DataFrame: DataFrame with percentages for each day and time slot.
    """
    # Create a dictionary to count the number of scheduled classes for each (day, time slot)
    schedule_counts = {(day, slot): 0 for day in days for slot in time_slots}
    
    total_classes = 0  # Track total number of classes

    # Iterate over the DataFrame to check which variables are active (scheduled)
    for _, row in df.iterrows():
        course = row['Course']
        instructor = row['Instructor']
        num_sections = int(row['# Sections'])
        
        if num_sections == 0:
            continue

        for section_id in range(1, num_sections + 1):
            parts = [1, 2, 3] if course == "CS435" else [1, 2]
            for part in parts:
                for day in days:
                    for slot in time_slots:
                        var_name = f"X_{course}_{instructor}_{section_id}_{part}_{day}_{slot}"
                        
                        # If the variable is scheduled (value is 1)
                        if var_name in variables and variables[var_name].X > 0.5:
                            # Increment the count for that (day, slot)
                            schedule_counts[(day, slot)] += 1
                            total_classes += 1

    # Create a DataFrame to store the percentages for each (day, slot)
    percentages = []
    for day, slot in schedule_counts:
        count = schedule_counts[(day, slot)]
        percentage = (count / total_classes) * 100 if total_classes > 0 else 0
        percentages.append({'Day': day, 'Time Slot': slot, 'Percentage': percentage})

    # Convert the list of percentages to a DataFrame
    percentage_df = pd.DataFrame(percentages)
    
    print(percentage_df)
    return percentage_df

def main():
    """
    Main function to run the entire scheduling process.
    """
    print("Starting scheduling system...")
    start_time = time.time()
    
    # Define the excel file path
    excel_file = "Scheduling Project Pilot.xlsx"
    
    # Load all data
    df, aggregated_df, df_pre_scheduled, df_constraints, general_preferences_df, section_capacity_map = load_data(excel_file)
    
    # Define time slots, days, and percentages
    time_slots, days = define_time_slots_and_days()
    slot_percentages = define_slot_percentages(days, time_slots)
    
    # Define course blocks
    course_blocks, special_blocks = define_course_blocks()
    
    # Define valid slots for course patterns
    valid_start_times, friday_start_times = define_valid_slots_for_course_patterns()
    
    # Define mapping for time slots abbreviations
    time_slot_mapping = {
        'M': 'Monday', 
        'T': 'Tuesday', 
        'W': 'Wednesday', 
        'R': 'Thursday', 
        'F': 'Friday',
        'S': 'Saturday'
    }
    
    # Time slot indexes mapping
    time_slot_index = {
        '1': "8:30-10:00 AM",
        '2': "10:00-11:30 AM",
        '3': "11:30-1:00 PM",
        '4': "1:00-2:30 PM",
        '5': "2:30-4:00 PM",
        '6': "4:00-5:30 PM",
        '7': "6:00-7:30 PM",
        '8': "7:30-9:00 PM"
    }
    
    # Initialize model and variables
    model, variables = initialize_model()
    
    # Create decision variables
    variables = create_decision_variables(model, variables, df, days, time_slots)
    
    # Add constraints
    add_unique_assignment_constraints(model, variables, df, days, time_slots)
    add_instructor_availability_constraints(model, variables, df, days, time_slots)
    add_time_slot_balance_constraints(model, variables, df, days, time_slots, slot_percentages)
    add_restricted_time_slots_constraints(model, variables, df, days)
    add_course_pattern_constraints(model, variables, df, days, time_slots, valid_start_times, friday_start_times)
    add_course_block_constraints(model, variables, df, days, time_slots, course_blocks, special_blocks)
    add_evening_constraints(model, variables, df, days)
    add_consecutive_slots_constraints(model, variables, df, days, time_slots)
    add_restricted_monday_constraints(model, variables, df, section_capacity_map)
    add_pre_scheduled_constraints(model, variables, df_pre_scheduled)
    
    # Define objective function components
    total_points = LinExpr()  # Initialize total points
    
    # Add health and religion constraints (high-penalty soft constraints)
    total_points = add_health_religion_constraints(model, variables, df, df_constraints, 
                                                 time_slot_mapping, time_slot_index, total_points)
    
    # Add instructor preference constraints
    total_points, instructor_soft_violated = add_instructor_preference_constraints(model, variables, df, 
                                                                                df_constraints, time_slot_mapping, 
                                                                                time_slot_index, total_points)
    
    # Add teaching days tracking variables
    z_vars = add_teaching_days_variables(model, variables, df, days, time_slots)
    
    # Add day preference penalties
    day_penalty_sum = add_day_preference_penalties(model, z_vars, general_preferences_df, df)
    
    # Get consecutive preferences from general preferences
    consecutive_preference = general_preferences_df.set_index('Email')['Consecutive Preference'].to_dict()
    
    # Add consecutive slot preference penalties
    consecutive_penalty_sum = add_consecutive_preference_penalties(model, variables, df, 
                                                                consecutive_preference, days, time_slots)
    
    # Add format preference penalties
    format_penalty_sum = add_format_preference_penalties(model, variables, df, 
                                                      general_preferences_df, days, time_slots)
    
    # Build objective function
    build_objective_function(model, total_points, format_penalty_sum, day_penalty_sum, consecutive_penalty_sum)
    
    # Solve the model
    model_status = solve_model(model, time_limit=600)
    
    # Analyze the solution
    schedule, instructor_day_counts = analyze_solution(model, variables, df, days, time_slots, 
                                                     section_capacity_map, instructor_soft_violated)
    
    # Generate output files
    if schedule:
        generate_output_files(schedule, instructor_day_counts, section_capacity_map)
    
    total_time = time.time() - start_time
    print(f"Total execution time: {total_time:.2f} seconds")

if __name__ == "__main__":
    main()
