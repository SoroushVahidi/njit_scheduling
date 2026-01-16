import os
import pandas as pd
import math
import copy
import time
from collections import defaultdict
from datetime import datetime
import gurobipy as gp
from gurobipy import GRB, Model, LinExpr

def main():
    # Get current date and time for file naming
    now = datetime.now()
    date_time_str = now.strftime("%Y-%m-%d_%H-%M-%S")
    
    print("=" * 60)
    print("COURSE SCHEDULING OPTIMIZATION SYSTEM")
    print("=" * 60)
    
    # File check
    excel_file = "Scheduling Project Pilot.xlsx"
    if os.path.exists(excel_file):
        print(f"✅ Found input file: {excel_file}")
    else:
        print(f"❌ Missing input file: {excel_file}")
        print(f"Current directory: {os.getcwd()}")
        return

    # Test Gurobi license
    print("🔧 Testing Gurobi license...")
    try:
        test_model = gp.Model("test")
        x = test_model.addVar(name="x")
        test_model.setObjective(x, GRB.MAXIMIZE)
        test_model.optimize()
        print("✅ Gurobi license is working!")
    except Exception as e:
        print(f"❌ Gurobi error: {e}")
        return

    print("\n📊 Loading and processing data...")
    start_time = time.time()
    
    # Load and process data
    df, section_capacity_map, section_type_map = load_and_process_data(excel_file)
    
    # Define time slots and days
    time_slots = [
        "8:30-10:00 AM", "10:00-11:30 AM", "11:30-1:00 PM", "1:00-2:30 PM",
        "2:30-4:00 PM", "4:00-5:30 PM", "6:00-7:30 PM", "7:30-9:00 PM"
    ]
    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
    
    print(f"📈 Total courses: {len(df)}")
    print(f"📈 Total sections: {df['# Sections'].sum()}")
    print(f"📈 Total instructors: {df['Instructor'].nunique()}")
    
    print("\n🏗️ Building optimization model...")
    model, variables = build_model(df, excel_file, time_slots, days, section_capacity_map, section_type_map)
    
    model_time = time.time() - start_time
    print(f"✅ Model built in {model_time:.2f} seconds")
    
    print("\n🚀 Starting optimization (3-hour limit)...")
    print("-" * 40)
    
    # Solve the model
    start_optimize = time.time()
    model.setParam("TimeLimit", 10800)
    model.optimize()
    
    optimize_time = time.time() - start_optimize
    print("-" * 40)
    print(f"✅ Optimization completed in {optimize_time:.2f} seconds")
    
    # Check solution status
    if model.Status == GRB.INFEASIBLE:
        print("❌ Model is infeasible")
        write_infeasible_analysis(model, date_time_str)
        return
    elif model.Status == GRB.TIME_LIMIT:
        if model.SolCount > 0 and model.ObjVal != 0:
            gap = (model.ObjBound - model.ObjVal) / abs(model.ObjVal) * 100
            print(f"⚠️ Time limit reached - Best solution found with {gap:.3f}% gap")
        else:
            print("⚠️ Time limit reached - no feasible solution found")
    else:
        print("✅ Optimal solution found")
    
    if model.SolCount == 0:
        print("❌ No feasible solution to report.")
        return
    
    print(f"📊 Objective value: {model.ObjVal:.0f}")
    
    # Long, detailed penalty analysis (kept but not called by default)
    # report_all_penalties(model, df, excel_file, date_time_str)
    
    # Short audit (health/religion & generic soft penalties, no explicit “preference” wording)
    write_short_audit(model, date_time_str)
    
    print("\n📝 Generating output files...")
    
    # Generate all outputs
    write_schedule_files(model, df, section_capacity_map, section_type_map, date_time_str, time_slots, days)
    write_impact_analysis(model, df, excel_file, date_time_str)
    write_constraint_violations(model, date_time_str)
    write_percentages_analysis(model, df, variables, time_slots, days, date_time_str)
    
    total_time = time.time() - start_time
    print(f"\n✅ All files generated successfully!")
    print(f"⏱️ Total runtime: {total_time:.2f} seconds")
    print("\nOutput files:")
    print(f"  📄 final_schedule_sorted_by_course_{date_time_str}.txt")
    print(f"  📄 final_schedule_sorted_by_instructor_{date_time_str}.txt") 
    print(f"  📄 impact_analysis_{date_time_str}.txt")
    print(f"  📄 constraint_violations_{date_time_str}.txt")
    print(f"  📄 scheduling_percentages_{date_time_str}.csv")
    print(f"  📄 instructors_days_analysis_{date_time_str}.txt")
    print(f"  📄 short_audit_{date_time_str}.txt")
    print(f"  📄 complete_penalty_analysis_{date_time_str}.txt (if you uncomment its call)")

def load_and_process_data(excel_file):
    """Load and process all data from Excel file"""
    # Load assignments
    df = pd.read_excel(excel_file, sheet_name='Assignments')
    df = df.rename(columns={
        df.columns[0]: 'Course',
        df.columns[1]: 'Instructor', 
        df.columns[2]: 'Capacity',
        df.columns[3]: '# Sections'
    })
    # Column F (index 5) = section type (e.g., Jersey City, etc.)
    if len(df.columns) > 5:
        df = df.rename(columns={df.columns[5]: 'Section_Type'})
    # Remove last 4 columns (as in original code)
    df = df.iloc[:, :-4]
    
    # Clean and process
    df = df.dropna(subset=['Course'])
    df['Course_Number'] = df['Course'].str.extract(r'(\d+)')
    
    # Load faculty data
    faculty_df = pd.read_excel(excel_file, sheet_name='Faculty')
    faculty_df = faculty_df.rename(columns={faculty_df.columns[0]: 'Instructor'})
    
    # Merge with faculty data (for Email, etc.)
    df = pd.merge(df, faculty_df, on='Instructor', how='left')
    
    # Create section maps BEFORE aggregating
    section_capacity_map = create_section_capacity_map(df)
    section_type_map = create_section_type_map(df)
    
    # Create aggregated dataframe (now we can safely remove Capacity / Section_Type)
    df = df.groupby(['Course', 'Instructor', 'Course_Number', 'Email']).agg({
        '# Sections': 'sum'
    }).reset_index()
    
    return df, section_capacity_map, section_type_map

def create_section_capacity_map(df):
    """Create mapping of (course, instructor, section) to capacity"""
    df_sorted = df.sort_values(by=['Course', 'Instructor', 'Capacity'])
    section_capacity_map = {}
    
    for (course, instructor), group in df_sorted.groupby(['Course', 'Instructor']):
        section_number = 1
        for _, row in group.iterrows():
            num_sections = int(row['# Sections'])
            capacity = row['Capacity']
            course_str = str(course).strip()
            instructor_str = str(instructor).strip()
            
            for _ in range(num_sections):
                section_capacity_map[(course_str, instructor_str, section_number)] = capacity
                section_number += 1
    
    return section_capacity_map

def create_section_type_map(df):
    """Create mapping of (course, instructor, section) to section type (e.g., Jersey City)"""
    section_type_map = {}
    if 'Section_Type' not in df.columns:
        return section_type_map
    
    df_sorted = df.sort_values(by=['Course', 'Instructor'])
    for (course, instructor), group in df_sorted.groupby(['Course', 'Instructor']):
        section_number = 1
        for _, row in group.iterrows():
            num_sections = int(row['# Sections'])
            stype = row.get('Section_Type', "")
            course_str = str(course).strip()
            instructor_str = str(instructor).strip()
            for _ in range(num_sections):
                section_type_map[(course_str, instructor_str, section_number)] = stype
                section_number += 1
    return section_type_map

def build_model(df, excel_file, time_slots, days, section_capacity_map, section_type_map):
    """Build the complete optimization model"""
    model = Model("Scheduling")
    variables = {}
    
    # Create binary variables - Always use 2 parts for all courses
    for _, row in df.iterrows():
        course = row['Course']
        instructor = row['Instructor']
        num_sections = int(row['# Sections'])
        
        if num_sections == 0:
            continue
            
        parts = [1, 2]  # Always use 2 parts for all courses
        
        for section_id in range(1, num_sections + 1):
            for part in parts:
                for day in days:
                    for slot in time_slots:
                        var_name = f"X_{course}_{instructor}_{section_id}_{part}_{day}_{slot}"
                        variables[var_name] = model.addVar(vtype=GRB.BINARY, name=var_name)
    
    # Add all constraints
    add_basic_constraints(model, df, variables, time_slots, days)
    add_balance_constraints(model, df, variables, time_slots, days, excel_file, section_type_map)
    add_pattern_constraints(model, df, variables, time_slots, days)
    add_course_block_constraints(model, df, variables, time_slots, days)
    add_consecutive_slots_constraints(model, df, variables, time_slots, days)
    add_jersey_city_constraints(model, df, variables, time_slots, days, section_type_map)
    
    # Add preference constraints and set objective (pass section_capacity_map)
    add_preference_constraints(model, df, variables, time_slots, days, excel_file, section_capacity_map)
    
    return model, variables

def add_basic_constraints(model, df, variables, time_slots, days):
    """Add basic scheduling constraints"""
    # Each section part must be scheduled exactly once
    for _, row in df.iterrows():
        course = row['Course']
        instructor = row['Instructor']
        num_sections = int(row['# Sections'])
        
        if num_sections == 0:
            continue
            
        parts = [1, 2]  # Always use 2 parts
        
        for section_id in range(1, num_sections + 1):
            for part in parts:
                section_vars = []
                for day in days:
                    for slot in time_slots:
                        var_name = f"X_{course}_{instructor}_{section_id}_{part}_{day}_{slot}"
                        if var_name in variables:
                            section_vars.append(variables[var_name])
                
                if section_vars:
                    model.addConstr(sum(section_vars) == 1, 
                                    name=f"unique_slot_{course}_{instructor}_{section_id}_{part}")
    
    # Instructor conflict constraints: at most one section part per instructor/day/slot
    for instructor in df['Instructor'].unique():
        for day in days:
            for slot in time_slots:
                instructor_vars = []
                for _, row in df[df['Instructor'] == instructor].iterrows():
                    course = row['Course']
                    num_sections = int(row['# Sections'])
                    parts = [1, 2]
                    
                    for section_id in range(1, num_sections + 1):
                        for part in parts:
                            var_name = f"X_{course}_{instructor}_{section_id}_{part}_{day}_{slot}"
                            if var_name in variables:
                                instructor_vars.append(variables[var_name])
                
                if instructor_vars:
                    model.addConstr(sum(instructor_vars) <= 1,
                                    name=f"one_section_per_slot_{instructor}_{day}_{slot}")

def add_balance_constraints(model, df, variables, time_slots, days, excel_file, section_type_map):
    """Add time slot balance constraints - EXPLICIT BAN for common hours, excluding Jersey City from counts"""
    # Helper: identify Jersey City variables
    def is_jc_var(var_name):
        """Return True if this X_ variable corresponds to a Jersey City section."""
        parts = var_name.split('_')
        if len(parts) < 7:
            return False
        course = parts[1]
        instructor = parts[2]
        try:
            section_id = int(parts[3])
        except ValueError:
            return False
        stype = section_type_map.get((course, instructor, section_id), "")
        return str(stype).strip().lower() == "jersey city"
    
    # Load pre-scheduled courses
    df_pre_scheduled = pd.read_excel(excel_file, sheet_name='pre-scheduled')
    pre_scheduled_counts = df_pre_scheduled.groupby(['Day', 'Time']).size().to_dict()
    
    # Enforce pre-scheduled courses - block those instructor/day/time combinations
    for _, row in df_pre_scheduled.iterrows():
        instructor = row['Instructor']
        day = row['Day']
        time_slot = row['Time']
        
        # Set all variables for this (instructor, day, time) to 0
        for var_name in variables:
            var_parts = var_name.split('_')
            if len(var_parts) >= 7:
                var_instructor = var_parts[2]
                var_day = var_parts[5]
                var_time = var_parts[6]
                
                # Check if the variable matches the instructor, day, and time
                if var_instructor == instructor and var_day == day and var_time == time_slot:
                    model.addConstr(variables[var_name] == 0, 
                                    name=f"block_{instructor}_{day}_{time_slot}")
    
    # Calculate total section parts
    total_section_parts = 2 * sum(int(row['# Sections']) for _, row in df.iterrows())
    
    # Define banned slots explicitly
    banned_slots = [
        ("Friday", "11:30-1:00 PM"),
        ("Wednesday", "2:30-4:00 PM"), 
        ("Wednesday", "4:00-5:30 PM")
    ]
    
    # EXPLICIT BAN: Force banned slots to have exactly 0 classes
    for day, slot in banned_slots:
        slot_vars = [
            variables[var_name] for var_name in variables 
            if len(var_name.split('_')) >= 7
            and var_name.split('_')[5] == day
            and var_name.split('_')[6] == slot
        ]
        
        if slot_vars:
            model.addConstr(sum(slot_vars) == 0,  # Force exactly 0
                            name=f"BANNED_SLOT_{day}_{slot}")
            print(f"BANNED: {day} {slot} - {len(slot_vars)} variables forced to 0")
    
    # Complete balance percentages
    slot_percentages = {
        ("Monday", "8:30-10:00 AM"): 0.20, ("Tuesday", "8:30-10:00 AM"): 0.20,
        ("Wednesday", "8:30-10:00 AM"): 0.25, ("Thursday", "8:30-10:00 AM"): 0.20,
        ("Friday", "8:30-10:00 AM"): 0.20,
        
        ("Monday", "10:00-11:30 AM"): 0.20, ("Tuesday", "10:00-11:30 AM"): 0.20,
        ("Wednesday", "10:00-11:30 AM"): 0.25, ("Thursday", "10:00-11:30 AM"): 0.20,
        ("Friday", "10:00-11:30 AM"): 0.20,
        
        ("Monday", "11:30-1:00 PM"): 0.20, ("Tuesday", "11:30-1:00 PM"): 0.20,
        ("Wednesday", "11:30-1:00 PM"): 0.25, ("Thursday", "11:30-1:00 PM"): 0.20,
        ("Friday", "11:30-1:00 PM"): 0.00,  # banned explicitly above
        
        ("Monday", "1:00-2:30 PM"): 0.20, ("Tuesday", "1:00-2:30 PM"): 0.20,
        ("Wednesday", "1:00-2:30 PM"): 0.25, ("Thursday", "1:00-2:30 PM"): 0.20,
        ("Friday", "1:00-2:30 PM"): 0.20,
        
        ("Monday", "2:30-4:00 PM"): 0.20, ("Tuesday", "2:30-4:00 PM"): 0.20,
        ("Wednesday", "2:30-4:00 PM"): 0.00, ("Thursday", "2:30-4:00 PM"): 0.20,
        ("Friday", "2:30-4:00 PM"): 0.20,
        
        ("Monday", "4:00-5:30 PM"): 0.20, ("Tuesday", "4:00-5:30 PM"): 0.20,
        ("Wednesday", "4:00-5:30 PM"): 0.00, ("Thursday", "4:00-5:30 PM"): 0.20,
        ("Friday", "4:00-5:30 PM"): 0.20,
        
        ("Monday", "6:00-7:30 PM"): 0.20, ("Tuesday", "6:00-7:30 PM"): 0.20,
        ("Wednesday", "6:00-7:30 PM"): 0.20, ("Thursday", "6:00-7:30 PM"): 0.20,
        ("Friday", "6:00-7:30 PM"): 0.20,
        
        ("Monday", "7:30-9:00 PM"): 0.20, ("Tuesday", "7:30-9:00 PM"): 0.20,
        ("Wednesday", "7:30-9:00 PM"): 0.20, ("Thursday", "7:30-9:00 PM"): 0.20,
        ("Friday", "7:30-9:00 PM"): 0.20
    }
    
    # Add balance constraints for non-banned slots only
    for (day, slot), percentage in slot_percentages.items():
        # Skip banned slots - they're already handled above
        if (day, slot) in banned_slots:
            continue
            
        # EXCLUDE Jersey City sections from balancing counts
        slot_vars = [
            variables[var_name] for var_name in variables 
            if (len(var_name.split('_')) >= 7
                and var_name.split('_')[5] == day
                and var_name.split('_')[6] == slot
                and not is_jc_var(var_name))
        ]
        
        max_section_parts_slot = math.ceil((percentage / 6) * total_section_parts)
        
        if slot_vars:
            model.addConstr(sum(slot_vars) <= max_section_parts_slot,
                            name=f"balance_slot_{day}_{slot}")
    
    # Evening slot constraints (also EXCLUDE Jersey City sections)
    evening_percentage = 0.20
    evening_slots = ["6:00-7:30 PM", "7:30-9:00 PM"]
    for slot in evening_slots:
        evening_vars = [
            variables[var_name] for var_name in variables 
            if (len(var_name.split('_')) >= 7
                and var_name.split('_')[6] == slot
                and not is_jc_var(var_name))
        ]
        
        max_section_parts_evening = (evening_percentage / 6) * total_section_parts
        if evening_vars:
            model.addConstr(sum(evening_vars) <= max_section_parts_evening, 
                            name=f"balance_evening_{slot}")

def add_pattern_constraints(model, df, variables, time_slots, days):
    """Add graduate/undergraduate pattern constraints"""
    valid_start_times = ["8:30-10:00 AM", "6:00-7:30 PM"]
    friday_start_times = ["8:30-10:00 AM", "1:00-2:30 PM", "2:30-4:00 PM", 
                          "4:00-5:30 PM", "6:00-7:30 PM", "7:30-9:00 PM"]
    
    for course in df['Course'].unique():
        for instructor in df['Instructor'].unique():
            course_instructor_rows = df[(df['Course'] == course) & (df['Instructor'] == instructor)]
            if course_instructor_rows.empty:
                continue
            
            for section_id in range(1, course_instructor_rows.iloc[0]['# Sections'] + 1):
                grad_var = model.addVar(vtype=GRB.BINARY, name=f"Grad_{course}_{instructor}_{section_id}")
                undergrad_var = model.addVar(vtype=GRB.BINARY, name=f"Undergrad_{course}_{instructor}_{section_id}")
                
                y_var_dict = {}
                compatible_pairs = []
                
                # Create Y variables for valid patterns
                for day1 in days:
                    for slot1 in time_slots:
                        for day2 in days:
                            for slot2 in time_slots:
                                y_var_name = f"Y_{course}_{instructor}_{section_id}_{day1}_{slot1}_{day2}_{slot2}"
                                y_var = model.addVar(vtype=GRB.BINARY, name=y_var_name)
                                compatible_pairs.append(y_var)
                                y_var_dict[(day1, slot1, day2, slot2)] = y_var
                                
                                # Graduate pattern (consecutive slots on same day)
                                if day1 == day2:
                                    if day1 == "Friday":
                                        if slot1 in friday_start_times and time_slots.index(slot2) == time_slots.index(slot1) + 1:
                                            model.addConstr(y_var <= grad_var)
                                        else:
                                            model.addConstr(y_var == 0)
                                    else:
                                        if slot1 in valid_start_times and time_slots.index(slot2) == time_slots.index(slot1) + 1:
                                            model.addConstr(y_var <= grad_var)
                                        else:
                                            model.addConstr(y_var == 0)
                                
                                # Undergraduate pattern (same slot, different valid day pairs)
                                elif day1 != day2 and slot1 == slot2:
                                    if ((day1 == "Monday" and day2 in ["Wednesday", "Thursday"]) or
                                        (day1 == "Tuesday" and day2 in ["Thursday", "Friday"]) or
                                        (day1 == "Wednesday" and day2 == "Friday")):
                                        model.addConstr(y_var <= undergrad_var)
                                    else:
                                        model.addConstr(y_var == 0)
                                else:
                                    model.addConstr(y_var == 0)
                
                # Special constraint: Courses with 3rd digit = 7 cannot have Part 1 at 8:30-10:00 AM
                if len(course) >= 3 and course[2] == "7":
                    for day1 in days:
                        var_part1 = f"X_{course}_{instructor}_{section_id}_1_{day1}_8:30-10:00 AM"
                        if var_part1 in variables:
                            model.addConstr(variables[var_part1] == 0,
                                            name=f"no_8_30_to_10_CS7XX_{course}_{instructor}_{section_id}_{day1}")
                
                # Link constraints
                for day1 in days:
                    for slot1 in time_slots:
                        var_part1 = f"X_{course}_{instructor}_{section_id}_1_{day1}_{slot1}"
                        if var_part1 in variables:
                            y_vars_for_part1 = [y_var_dict[(day1, slot1, day2, slot2)] 
                                                for day2 in days for slot2 in time_slots]
                            model.addConstr(sum(y_vars_for_part1) == variables[var_part1],
                                            name=f"part1_link_{course}_{instructor}_{section_id}_{day1}_{slot1}")
                
                for day2 in days:
                    for slot2 in time_slots:
                        var_part2 = f"X_{course}_{instructor}_{section_id}_2_{day2}_{slot2}"
                        if var_part2 in variables:
                            y_vars_for_part2 = [y_var_dict[(day1, slot1, day2, slot2)] 
                                                for day1 in days for slot1 in time_slots]
                            model.addConstr(sum(y_vars_for_part2) == variables[var_part2],
                                            name=f"part2_link_{course}_{instructor}_{section_id}_{day2}_{slot2}")
                
                # Exactly one pair and one pattern
                model.addConstr(sum(compatible_pairs) == 1,
                                name=f"select_one_pair_{course}_{instructor}_{section_id}")
                model.addConstr(grad_var + undergrad_var == 1,
                                name=f"select_one_pattern_{course}_{instructor}_{section_id}")

def add_course_block_constraints(model, df, variables, time_slots, days):
    """Add course block conflict constraints - CORRECTED to match notebook exactly"""
    course_blocks = [
        ['CS114', 'IS210', 'CS450', 'CS337'],
        ['CS241', 'CS280', 'IS350'],
        ['CS288', 'CS332', 'CS301', 'CS356'],
        ['CS341', 'CS350', 'CS351', 'CS331', 'CS375'],
        ['CS435', 'CS490', 'CS485', 'CS370', 'CS375'],
        ['CS485', 'CS491', 'CS450', 'CS482'],
        ['CS610', 'CS630', 'CS631', 'CS656', 'DS675', 'CS675', 'CS670'],
        ['DS677', 'DS669', 'DS650', 'CS670', 'CS610', 'CS665', 'CS667', 'CS732', 'DS680'],
        ['CS608', 'CS645', 'CS646', 'CS647', 'CS648', 'CS678', 'CS696'],
        ['IS455', 'IS645'],
        ['IT220', 'IT230', 'IT240', 'IT302'],
        ['IT256', 'IT266', 'IT286', 'IT360', 'IT380', 'IT383', 'IT386'],
        ['IT120', 'IT240']
    ]
    
    special_blocks = [
        ['CS288', 'CS332', 'CS301', 'CS356'],
        ['CS341', 'CS350', 'CS351', 'CS331', 'CS375']
    ]
    
    for block in course_blocks:
        for course1 in block:
            for course2 in block:
                if course1 != course2:
                    for instructor1 in df[df['Course'] == course1]['Instructor'].unique():
                        for instructor2 in df[df['Course'] == course2]['Instructor'].unique():
                            for day in days:
                                for slot in time_slots:
                                    var_course1_part1 = f"X_{course1}_{instructor1}_1_{day}_{slot}"
                                    var_course1_part2 = f"X_{course1}_{instructor1}_2_{day}_{slot}"
                                    var_course2_part1 = f"X_{course2}_{instructor2}_1_{day}_{slot}"
                                    var_course2_part2 = f"X_{course2}_{instructor2}_2_{day}_{slot}"

                                    max_constraint = 2 if block in special_blocks else 1

                                    if var_course1_part1 in variables and var_course2_part1 in variables:
                                        model.addConstr(
                                            variables[var_course1_part1] + variables[var_course2_part1] <= max_constraint,
                                            name=f"no_same_day_slot_{course1}_{instructor1}_{course2}_{instructor2}_{day}_{slot}_part1"
                                        )
                                    if var_course1_part2 in variables and var_course2_part2 in variables:
                                        model.addConstr(
                                            variables[var_course1_part2] + variables[var_course2_part2] <= max_constraint,
                                            name=f"no_same_day_slot_{course1}_{instructor1}_{course2}_{instructor2}_{day}_{slot}_part2"
                                        )

def add_consecutive_slots_constraints(model, df, variables, time_slots, days):
    """Add consecutive slots constraints - CORRECTED to match notebook exactly"""
    for instructor in df['Instructor'].unique():
        for day in days:
            for i in range(len(time_slots) - 2):
                slot1 = time_slots[i]
                slot2 = time_slots[i + 1]
                slot3 = time_slots[i + 2]
                
                instructor_df = df[df['Instructor'] == instructor]

                for _, row1 in instructor_df.iterrows():
                    course1 = row1['Course']
                    num_sections1 = int(row1['# Sections'])
                    
                    if num_sections1 == 0:
                        continue

                    for section_id1 in range(1, num_sections1 + 1):
                        parts = [1, 2]
                        for part1 in parts:
                            var_name1 = f"X_{course1}_{instructor}_{section_id1}_{part1}_{day}_{slot1}"

                            for _, row2 in instructor_df.iterrows():
                                course2 = row2['Course']
                                num_sections2 = int(row2['# Sections'])
                                
                                if num_sections2 == 0:
                                    continue

                                for section_id2 in range(1, num_sections2 + 1):
                                    for part2 in parts:
                                        var_name2 = f"X_{course2}_{instructor}_{section_id2}_{part2}_{day}_{slot2}"

                                        if (course1 != course2 or section_id1 != section_id2 or part1 != part2):
                                            
                                            for _, row3 in instructor_df.iterrows():
                                                course3 = row3['Course']
                                                num_sections3 = int(row3['# Sections'])
                                                
                                                if num_sections3 == 0:
                                                    continue

                                                for section_id3 in range(1, num_sections3 + 1):
                                                    for part3 in parts:
                                                        var_name3 = f"X_{course3}_{instructor}_{section_id3}_{part3}_{day}_{slot3}"

                                                        if ((course1 != course3 or section_id1 != section_id3 or part1 != part3) and
                                                            (course2 != course3 or section_id2 != section_id3 or part2 != part3)):
                                                            consecutive_sum = 0

                                                            if var_name1 in variables:
                                                                consecutive_sum += variables[var_name1]
                                                            if var_name2 in variables:
                                                                consecutive_sum += variables[var_name2]
                                                            if var_name3 in variables:
                                                                consecutive_sum += variables[var_name3]

                                                            model.addConstr(consecutive_sum <= 2, 
                                                                            name=f"consecutive_slots_constraint_{instructor}_{day}_{slot1}_{slot2}_{slot3}")

def add_jersey_city_constraints(model, df, variables, time_slots, days, section_type_map):
    """
    Special constraints for Jersey City sections:
    - Only allow:
        Part 1: 6:00-7:30 PM
        Part 2: 7:30-9:00 PM
    - Do not schedule more than 3 Jersey City sections per day.
    - Travel buffer: if instructor has a Jersey City section on a day in 6:00-7:30 PM,
      they cannot teach any other (non-JC) section in 4:00-5:30 PM on the same day.
    """
    jc_type_str = "jersey city"
    evening_start = "6:00-7:30 PM"
    evening_next  = "7:30-9:00 PM"
    prev_slot     = "4:00-5:30 PM"
    
    # 1. Force JC sections to evening 3-hour pattern
    for (course, instructor, section_id), stype in section_type_map.items():
        if str(stype).strip().lower() != jc_type_str:
            continue
        
        for day in days:
            for slot in time_slots:
                # Part 1 allowed only at 6:00-7:30 PM
                var1_name = f"X_{course}_{instructor}_{section_id}_1_{day}_{slot}"
                if var1_name in variables:
                    if slot != evening_start:
                        model.addConstr(variables[var1_name] == 0,
                                        name=f"JC_part1_fix_{course}_{instructor}_{section_id}_{day}_{slot}")
                # Part 2 allowed only at 7:30-9:00 PM
                var2_name = f"X_{course}_{instructor}_{section_id}_2_{day}_{slot}"
                if var2_name in variables:
                    if slot != evening_next:
                        model.addConstr(variables[var2_name] == 0,
                                        name=f"JC_part2_fix_{course}_{instructor}_{section_id}_{day}_{slot}")
    
    # 2. No more than 3 Jersey City sections per day
    for day in days:
        jc_day_vars = []
        for (course, instructor, section_id), stype in section_type_map.items():
            if str(stype).strip().lower() != jc_type_str:
                continue
            var1_name = f"X_{course}_{instructor}_{section_id}_1_{day}_{evening_start}"
            if var1_name in variables:
                jc_day_vars.append(variables[var1_name])
        if jc_day_vars:
            model.addConstr(gp.quicksum(jc_day_vars) <= 3,
                            name=f"JC_max3_per_day_{day}")
    
    # 3. Travel buffer: if instructor has JC at 6–7:30, the 4–5:30 slot must be free of non-JC
    for instructor in df['Instructor'].unique():
        for day in days:
            jc_vars = []
            other_prevslot_vars = []
            
            # Collect JC part-1 evening variables and non-JC previous-slot variables
            for (course, instr_key, section_id), stype in section_type_map.items():
                if instr_key != instructor:
                    continue
                is_jc = str(stype).strip().lower() == jc_type_str
                
                if is_jc:
                    var_jc = f"X_{course}_{instructor}_{section_id}_1_{day}_{evening_start}"
                    if var_jc in variables:
                        jc_vars.append(var_jc)
                else:
                    for part in [1, 2]:
                        var_other = f"X_{course}_{instructor}_{section_id}_{part}_{day}_{prev_slot}"
                        if var_other in variables:
                            other_prevslot_vars.append(var_other)
            
            # For every JC variable and every non-JC previous-slot variable, disallow co-occurrence
            for v_jc_name in jc_vars:
                for v_other_name in other_prevslot_vars:
                    model.addConstr(
                        variables[v_jc_name] + variables[v_other_name] <= 1,
                        name=f"JC_travel_buffer_{instructor}_{day}"
                    )

def add_preference_constraints(model, df, variables, time_slots, days, excel_file, section_capacity_map):
    """Add preference constraints and penalties"""
    total_points = LinExpr()
    
    # Load constraints and preferences
    df_constraints = pd.read_excel(excel_file, sheet_name='Constraints & Preferences')
    
    # Time slot mapping
    time_slot_mapping = {'M': 'Monday', 'T': 'Tuesday', 'W': 'Wednesday', 'R': 'Thursday', 'F': 'Friday'}
    time_slot_index = {
        '1': "8:30-10:00 AM", '2': "10:00-11:30 AM", '3': "11:30-1:00 PM", '4': "1:00-2:30 PM",
        '5': "2:30-4:00 PM", '6': "4:00-5:30 PM", '7': "6:00-7:30 PM", '8': "7:30-9:00 PM"
    }
    
    # Process constraints - EXACTLY as in notebook
    for _, row in df_constraints.iterrows():
        instructor_info = row['Instructor UCID: Type']
        slots = row['Slots']
        if isinstance(instructor_info, float):
            break
        
        email, constraint_type = instructor_info.split(": ")
        
        if constraint_type.strip() in ["Health", "Religion"]:
            points = -2048
        elif constraint_type.strip() == "Pref-1":
            points = 8
        elif constraint_type.strip() == "Pref-2":
            points = 4
        elif constraint_type.strip() == "Pref-3":
            points = 2
        elif constraint_type.strip() == "Childcare":
            points = -1024
        else:
            points = -8
        
        blocked_slots = slots.split("|")[1:-1]
        for slot_code in blocked_slots:
            if len(slot_code) < 2:
                continue
            
            day_abbrev = slot_code[0]
            time_slot_num = slot_code[1]
            
            day_full = time_slot_mapping[day_abbrev]
            time_slot_full = time_slot_index[time_slot_num]
            
            instructor_row = df[df['Email'] == email]
            if instructor_row.empty:
                continue
            
            instructor_name = instructor_row['Instructor'].iloc[0]
            
            for course in df[df['Instructor'] == instructor_name]['Course']:
                filtered_df = df[df['Course'] == course]
                if not filtered_df.empty:
                    num_sections = int(filtered_df[
                        (filtered_df['Instructor'] == instructor_name) &
                        (filtered_df['Course'] == course)
                    ]['# Sections'].iloc[0])
                    
                    for section_id in range(1, num_sections + 1):
                        parts = [1, 2]
                        for part in parts:
                            var_name = f"X_{course}_{instructor_name}_{section_id}_{part}_{day_full}_{time_slot_full}"
                            if var_name in variables:
                                if constraint_type.strip() in ["Health", "Religion"]:
                                    slack_var_name = f"Slack_{course}_{instructor_name}_{section_id}_{part}_{day_full}_{time_slot_full}"
                                    slack_var = model.addVar(vtype=GRB.BINARY, name=slack_var_name)
                                    model.addConstr(variables[var_name] <= slack_var)
                                    total_points -= 2048 * slack_var
                                else:
                                    total_points += points * variables[var_name]
    
    # Add additional constraints
    total_points = add_additional_constraints(model, df, variables, time_slots, days, excel_file, total_points, section_capacity_map)
    
    # Set objective
    model.setObjective(total_points, GRB.MAXIMIZE)

def add_additional_constraints(model, df, variables, time_slots, days, excel_file, total_points, section_capacity_map):
    """Add all the additional constraints that were in the original code"""
    
    # 1. Restricted time slots constraint (max 3 of 4 restricted slots per instructor per day)
    restricted_time_slots = ["8:30-10:00 AM", "10:00-11:30 AM", "6:00-7:30 PM", "7:30-9:00 PM"]
    
    for instructor in df['Instructor'].unique():
        for day in days:
            restricted_vars = []
            for course in df['Course'].unique():
                instructor_courses = df[(df['Instructor'] == instructor) & (df['Course'] == course)]
                
                for _, course_row in instructor_courses.iterrows():
                    num_sections = int(course_row['# Sections'])
                    for section_id in range(1, num_sections + 1):
                        parts = [1, 2]
                        for part in parts:
                            for slot in restricted_time_slots:
                                var_name = f"X_{course}_{instructor}_{section_id}_{part}_{day}_{slot}"
                                if var_name in variables:
                                    restricted_vars.append(variables[var_name])
            
            if restricted_vars:
                model.addConstr(
                    gp.quicksum(restricted_vars) <= 3,
                    name=f"restricted_time_slots_{instructor}_{day}"
                )
    
    # 2. Monday 4:00-5:30 PM restriction (only courses > 199 with capacity < 35)
    restricted_day = "Monday"
    restricted_time_slot = "4:00-5:30 PM"
    
    for _, row in df.iterrows():
        course = row['Course']
        instructor = row['Instructor']
        course_number = int(row['Course_Number']) if pd.notna(row['Course_Number']) else 0
        
        for sc in range(1, int(row['# Sections']) + 1):
            capacity = section_capacity_map.get((course, instructor, sc))
            
            if not (course_number > 199 and capacity and capacity < 35):
                for section_id in range(1, int(row['# Sections']) + 1):
                    parts = [1, 2]
                    for part in parts:
                        var_name = f"X_{course}_{instructor}_{section_id}_{part}_{restricted_day}_{restricted_time_slot}"
                        if var_name in variables:
                            model.addConstr(variables[var_name] == 0,
                                            name=f"restricted_slot_{course}_{instructor}_{section_id}_{part}_{restricted_day}_{restricted_time_slot}")
    
    # 3. Evening timing constraint (6:00-7:30 PM and 7:30-9:00 PM must be together)
    part1_slot = "6:00-7:30 PM"
    part2_slot = "7:30-9:00 PM"
    
    for _, row in df.iterrows():
        course = row['Course']
        instructor = row['Instructor']
        num_sections = int(row['# Sections'])
        
        if num_sections == 0:
            continue
        
        for section_id in range(1, num_sections + 1):
            for day in days:
                part1_var_name = f"X_{course}_{instructor}_{section_id}_1_{day}_{part1_slot}"
                part2_var_name = f"X_{course}_{instructor}_{section_id}_2_{day}_{part2_slot}"
                
                if part1_var_name in variables and part2_var_name in variables:
                    model.addConstr(variables[part1_var_name] == variables[part2_var_name],
                                    name=f"timing_constraint_{course}_{instructor}_{section_id}_{day}")
    
    # 4. General Preferences - Format preferences
    try:
        general_preferences_df = pd.read_excel(excel_file, sheet_name="General Preferences")
        general_preferences_df = general_preferences_df.rename(columns={
            general_preferences_df.columns[1]: 'Email',
            general_preferences_df.columns[2]: 'Preference'
        })
        
        format_penalty_sum = LinExpr()
        penalty_value = -8
        
        for _, row in general_preferences_df.iterrows():
            email = row['Email']
            preference = row['Preference']
            
            if pd.isna(email) or pd.isna(preference):
                continue
            
            instructor_row = df[df['Email'] == email]
            if instructor_row.empty:
                continue
            
            instructor_name = instructor_row['Instructor'].iloc[0]
            instructor_courses = df[df['Instructor'] == instructor_name]
            
            for _, course_row in instructor_courses.iterrows():
                course = course_row['Course']
                num_sections = int(course_row['# Sections'])
                
                for section_id in range(1, num_sections + 1):
                    for day in days:
                        for i in range(len(time_slots) - 1):
                            slot1 = time_slots[i]
                            slot2 = time_slots[i + 1]
                            
                            part1_var_name = f"X_{course}_{instructor_name}_{section_id}_1_{day}_{slot1}"
                            part2_var_name = f"X_{course}_{instructor_name}_{section_id}_2_{day}_{slot2}"
                            
                            if part1_var_name in variables and part2_var_name in variables:
                                part1_var = variables[part1_var_name]
                                part2_var = variables[part2_var_name]
                                
                                penalty_var = model.addVar(vtype=GRB.BINARY,
                                                           name=f"Penalty_{instructor_name}_{day}_{slot1}_{slot2}")
                                
                                if preference == "3-hour format":
                                    model.addConstr(part1_var - part2_var <= penalty_var)
                                    model.addConstr(part2_var - part1_var <= penalty_var)
                                elif preference == "1.5+1.5 hour format":
                                    model.addConstr(part1_var + part2_var - penalty_var <= 1)
                                
                                format_penalty_sum += penalty_value * penalty_var
        
        # 5. Day preferences
        if len(general_preferences_df.columns) > 3:
            general_preferences_df = general_preferences_df.rename(columns={
                general_preferences_df.columns[3]: 'Day Preference'
            })
            day_preference_dict = general_preferences_df.set_index('Email')['Day Preference'].to_dict()
            
            z_vars = {}
            for instructor in df['Instructor'].unique():
                for day in days:
                    z_var = model.addVar(vtype=GRB.BINARY, name=f"Z_{instructor}_{day}")
                    z_vars[(instructor, day)] = z_var
                    
                    relevant_x_vars = []
                    for course in df['Course'].unique():
                        course_instructor_rows = df[(df['Course'] == course) & (df['Instructor'] == instructor)]
                        
                        if course_instructor_rows.empty:
                            continue
                        
                        num_sections = course_instructor_rows['# Sections'].iloc[0]
                        for section_id in range(1, int(num_sections) + 1):
                            for part in [1, 2]:
                                for slot in time_slots:
                                    x_var_name = f"X_{course}_{instructor}_{section_id}_{part}_{day}_{slot}"
                                    if x_var_name in variables:
                                        x_var = variables[x_var_name]
                                        relevant_x_vars.append(x_var)
                                        model.addConstr(x_var <= z_var)
                    
                    if not relevant_x_vars:
                        model.addConstr(z_var == 0)
            
            day_penalty_sum = LinExpr()
            for instructor in df['Instructor'].unique():
                instructor_rows = df[df['Instructor'] == instructor]
                if instructor_rows.empty:
                    continue
                
                email = instructor_rows['Email'].iloc[0]
                prefers_condensed_days = day_preference_dict.get(email, "No") == "I prefer to condense my sections into fewer days"
                penalty_value = -8 if prefers_condensed_days else -3
                
                for day in days:
                    z_var = z_vars[(instructor, day)]
                    day_penalty_sum += penalty_value * z_var
            
            total_points += day_penalty_sum
        
        # 6. Consecutive slots preferences (avoid consecutive if they say "No")
        if len(general_preferences_df.columns) > 5:
            general_preferences_df = general_preferences_df.rename(columns={
                general_preferences_df.columns[5]: 'Consecutive Preference'
            })
            consecutive_preference = general_preferences_df.set_index('Email')['Consecutive Preference'].to_dict()
            
            consecutive_penalty_sum = LinExpr()
            penalty_value = -2048
            
            for email, prefers_consecutive in consecutive_preference.items():
                if pd.isna(prefers_consecutive) or prefers_consecutive != "No":
                    continue
                
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
                                x_var1_name = f"X_{course}_{instructor_name}_{section_id}_1_{day}_{slot1}"
                                x_var2_name = f"X_{course}_{instructor_name}_{section_id}_2_{day}_{slot2}"
                                
                                if x_var1_name in variables and x_var2_name in variables:
                                    x_var1 = variables[x_var1_name]
                                    x_var2 = variables[x_var2_name]
                                    
                                    penalty_var = model.addVar(vtype=GRB.BINARY,
                                                               name=f"ConsecutivePenalty_{instructor_name}_{day}_{slot1}_{slot2}")
                                    model.addConstr(x_var1 + x_var2 - 2 * penalty_var <= 1)
                                    consecutive_penalty_sum += penalty_value * penalty_var
            
            total_points += consecutive_penalty_sum
        
        # Add format penalty sum to total points
        total_points += format_penalty_sum
        
    except Exception as e:
        print(f"Warning: Could not process General Preferences sheet: {e}")
    
    return total_points

def write_schedule_files(model, df, section_capacity_map, section_type_map, date_time_str, time_slots, days):
    """Write the main schedule files (by course and by instructor), including section type"""
    schedule = []
    
    for var in model.getVars():
        if var.x > 0.5 and var.varName.startswith('X_'):
            try:
                parts = var.varName.split('_')
                if len(parts) >= 7:
                    _, course, instructor, section_id, part, day = parts[:6]
                    slot = '_'.join(parts[6:])  # Handle time slots with underscores
                    
                    instructor_rows = df[df['Instructor'] == instructor]
                    if not instructor_rows.empty:
                        email = instructor_rows['Email'].iloc[0]
                    else:
                        email = ""
                    
                    # Look up section type
                    try:
                        sec_id_int = int(section_id)
                    except ValueError:
                        sec_id_int = None
                    stype = section_type_map.get((course.strip(), instructor.strip(), sec_id_int), "Unknown")
                    
                    schedule.append((course, instructor, email, section_id, part, day, slot, stype))
            except (ValueError, IndexError) as e:
                print(f"Warning: Could not parse variable {var.varName}: {e}")
                pass
    
    schedule.sort()
    
    # Write course-sorted schedule
    with open(f"final_schedule_sorted_by_course_{date_time_str}.txt", "w") as f:
        f.write("Course Schedule (Lexicographically Sorted):\n\n")
        course_section_tracker = {}
        
        for entry in schedule:
            course, instructor, email, section_id, part, day, slot, stype = entry
            course = course.strip()
            instructor = instructor.strip()
            
            if course not in course_section_tracker:
                course_section_tracker[course] = 1
            
            assigned_section_number = course_section_tracker[course]
            capacity = section_capacity_map.get((course, instructor, int(section_id)), "Unknown")
            
            f.write(
                f"Course: {course}, Instructor: {instructor}, Email: {email}, "
                f"Section: {assigned_section_number}, Part: {part}, Day: {day}, "
                f"Slot: {slot}, Capacity: {capacity}, Section Type: {stype}\n"
            )
            
            if part == "2":  # Always 2 parts now
                course_section_tracker[course] += 1
                f.write("\n")
    
    # Write instructor-sorted schedule
    schedule_sorted_by_instructor = sorted(schedule, key=lambda x: x[1])
    
    with open(f"final_schedule_sorted_by_instructor_{date_time_str}.txt", "w") as f:
        f.write("Course Schedule (Sorted by Instructor):\n\n")
        course_section_tracker = {}
        current_instructor = None
        
        for entry in schedule_sorted_by_instructor:
            course, instructor, email, section_id, part, day, slot, stype = entry
            
            if course not in course_section_tracker:
                course_section_tracker[course] = 1
            
            assigned_section_number = course_section_tracker[course]
            
            if instructor != current_instructor:
                if current_instructor is not None:
                    f.write("\n")
                f.write(f"Instructor: {instructor}, Email: {email}\n")
                current_instructor = instructor
            
            capacity = section_capacity_map.get((course, instructor, int(section_id)), "Unknown")
            f.write(
                f"\tCourse: {course}, Section: {assigned_section_number}, "
                f"Part: {part}, Day: {day}, Slot: {slot}, "
                f"Capacity: {capacity}, Section Type: {stype}\n"
            )
            
            if part == "2":  # Always 2 parts now
                course_section_tracker[course] += 1

def write_impact_analysis(model, df, excel_file, date_time_str):
    """Write detailed impact analysis to file"""
    with open(f"impact_analysis_{date_time_str}.txt", "w") as f:
        f.write("INSTRUCTOR IMPACT ANALYSIS\n")
        f.write("=" * 50 + "\n\n")
        
        # Analyze constraint violations
        instructor_impact = {}
        violated_vars = []
        
        for var in model.getVars():
            if var.varName.startswith('Slack_') and var.x > 0.5:
                try:
                    parts = var.varName.split('_')
                    if len(parts) >= 3:
                        instructor = parts[2]
                        violated_vars.append(var.varName)
                        
                        if instructor not in instructor_impact:
                            instructor_impact[instructor] = 0
                        instructor_impact[instructor] -= 2048  # Health/Religion penalty
                except:
                    pass
        
        f.write("HEALTH/RELIGION CONSTRAINT VIOLATIONS\n")
        f.write("-" * 40 + "\n")
        for var_name in violated_vars:
            f.write(f"{var_name}\n")
        f.write(f"\nTotal violated constraints: {len(violated_vars)}\n\n")
        
        f.write("INSTRUCTOR IMPACT SUMMARY\n")
        f.write("-" * 40 + "\n")
        if instructor_impact:
            sorted_impact = sorted(instructor_impact.items(), key=lambda x: x[1])
            for instructor, impact in sorted_impact:
                f.write(f"Instructor: {instructor}, Net Impact: {impact}\n")
        else:
            f.write("No instructor penalties found.\n")
        
        f.write("\n\nDETAILED CONSTRAINT VIOLATIONS\n")
        f.write("-" * 40 + "\n")
        
        # Check for format preference violations
        format_violations = []
        for var in model.getVars():
            if var.varName.startswith('Penalty_') and var.x > 0.5:
                format_violations.append(var.varName)
        
        if format_violations:
            f.write("Format/Soft Penalty Violations:\n")
            for violation in format_violations:
                f.write(f"  {violation}\n")
        else:
            f.write("No format/soft penalty violations found.\n")

def write_constraint_violations(model, date_time_str):
    """Write constraint violation summary"""
    with open(f"constraint_violations_{date_time_str}.txt", "w") as f:
        f.write("CONSTRAINT VIOLATIONS SUMMARY\n")
        f.write("=" * 40 + "\n\n")
        
        # Health/Religion violations
        health_violations = []
        for var in model.getVars():
            if var.varName.startswith('Slack_') and var.x > 0.5:
                health_violations.append((var.varName, var.x))
        
        f.write("HEALTH/RELIGION VIOLATIONS (-2048 points each):\n")
        f.write("-" * 40 + "\n")
        if health_violations:
            for var_name, value in health_violations:
                f.write(f"{var_name}: {value}\n")
        else:
            f.write("No health/religion constraint violations.\n")
        
        f.write(f"\nTotal violations: {len(health_violations)}\n")
        f.write(f"Total penalty: {len(health_violations) * -2048} points\n\n")
        
        # Other penalties
        other_penalties = []
        for var in model.getVars():
            if (var.varName.startswith('Penalty_') or var.varName.startswith('ConsecutivePenalty_')) and var.x > 0.5:
                other_penalties.append((var.varName, var.x))
        
        f.write("OTHER SOFT PENALTY VIOLATIONS:\n")
        f.write("-" * 40 + "\n")
        if other_penalties:
            for var_name, value in other_penalties:
                f.write(f"{var_name}: {value}\n")
        else:
            f.write("No other soft penalty violations.\n")

def write_percentages_analysis(model, df, variables, time_slots, days, date_time_str):
    """Calculate and write scheduling percentages"""
    schedule_counts = {(day, slot): 0 for day in days for slot in time_slots}
    total_classes = 0
    
    for _, row in df.iterrows():
        course = row['Course']
        instructor = row['Instructor']
        num_sections = int(row['# Sections'])
        
        if num_sections == 0:
            continue
        
        for section_id in range(1, num_sections + 1):
            parts = [1, 2]
            for part in parts:
                for day in days:
                    for slot in time_slots:
                        var_name = f"X_{course}_{instructor}_{section_id}_{part}_{day}_{slot}"
                        
                        if var_name in variables and variables[var_name].X > 0.5:
                            schedule_counts[(day, slot)] += 1
                            total_classes += 1
    
    percentages = []
    for day, slot in schedule_counts:
        count = schedule_counts[(day, slot)]
        percentage = (count / total_classes) * 100 if total_classes > 0 else 0
        percentages.append({'Day': day, 'Time Slot': slot, 'Count': count, 'Percentage': percentage})
    
    percentage_df = pd.DataFrame(percentages)
    percentage_df.to_csv(f"scheduling_percentages_{date_time_str}.csv", index=False)
    
    # Write instructor days analysis
    instructor_days = defaultdict(set)
    
    for var in model.getVars():
        if var.varName.startswith('X_') and var.x > 0.5:
            try:
                parts = var.varName.split('_')
                if len(parts) >= 6:
                    _, course, instructor, section_id, part, day = parts[:6]
                    instructor_days[instructor].add(day)
            except (ValueError, IndexError):
                pass
    
    instructor_day_counts = []
    for instructor, days_set in instructor_days.items():
        num_days = len(days_set)
        instructor_day_counts.append((instructor, num_days, sorted(days_set)))
    
    instructor_day_counts_sorted = sorted(instructor_day_counts, key=lambda x: x[1], reverse=True)
    
    with open(f"instructors_days_analysis_{date_time_str}.txt", "w") as f:
        f.write("INSTRUCTORS SORTED BY DAYS ON CAMPUS\n")
        f.write("=" * 50 + "\n\n")
        
        for instructor, num_days, days_list in instructor_day_counts_sorted:
            f.write(f"Instructor: {instructor}\n")
            f.write(f"  Days on campus: {num_days}\n")
            f.write(f"  Specific days: {', '.join(days_list)}\n\n")

def write_infeasible_analysis(model, date_time_str):
    """Handle infeasible model case"""
    print("❌ Computing Irreducible Inconsistent Subsystem (IIS)...")
    model.computeIIS()
    
    with open(f"infeasible_model_analysis_{date_time_str}.txt", "w") as f:
        f.write("INFEASIBLE MODEL ANALYSIS\n")
        f.write("=" * 40 + "\n\n")
        f.write("The following constraints are in conflict:\n\n")
        
        for i, constr in enumerate(model.getConstrs(), 1):
            if constr.IISConstr:
                f.write(f"Constraint {i}: {constr.ConstrName}\n")
                try:
                    f.write(f"{model.getRow(constr)} = {constr.RHS}\n\n")
                except:
                    f.write(f"RHS: {constr.RHS}\n\n")
    
    print(f"📄 IIS analysis written to infeasible_model_analysis_{date_time_str}.txt")

def write_short_audit(model, date_time_str):
    """
    Short audit:
    - Health/Religion slack (Slack_ variables)
    - Other generic soft penalties (Penalty_, ConsecutivePenalty_, Z_), but
      with no explicit wording about “preferences”.
    """
    filename = f"short_audit_{date_time_str}.txt"
    with open(filename, "w") as f:
        f.write("SHORT AUDIT REPORT\n")
        f.write("=" * 40 + "\n\n")
        
        # Health/Religion
        health_violations = [var for var in model.getVars() if var.varName.startswith('Slack_') and var.x > 0.5]
        f.write("HEALTH/RELIGION-RELATED SOFT CONSTRAINTS:\n")
        f.write("-" * 40 + "\n")
        if health_violations:
            for var in health_violations:
                f.write(f"{var.varName} = {var.x}\n")
        else:
            f.write("No active health/religion-related slack variables.\n")
        f.write(f"\nTotal: {len(health_violations)}\n\n")
        
        # Other soft penalties (generic)
        other_soft = [
            var for var in model.getVars()
            if (var.varName.startswith('Penalty_') or
                var.varName.startswith('ConsecutivePenalty_') or
                var.varName.startswith('Z_')) and var.x > 0.5
        ]
        f.write("OTHER SOFT PENALTY VARIABLES:\n")
        f.write("-" * 40 + "\n")
        if other_soft:
            for var in other_soft:
                f.write(f"{var.varName} = {var.x}\n")
        else:
            f.write("No active soft penalty variables.\n")
        f.write(f"\nTotal: {len(other_soft)}\n")
    
    print(f"📄 Short audit written to {filename}")

def report_all_penalties(model, df, excel_file, date_time_str=None):
    """Long, detailed penalty analysis (kept for reference; call is commented out in main)."""
    import pandas as pd
    
    if date_time_str is None:
        now = datetime.now()
        date_time_str = now.strftime("%Y-%m-%d_%H-%M-%S")
    
    output_lines = []
    
    def print_and_log(text):
        print(text)
        output_lines.append(text)
    
    print_and_log("\n" + "="*60)
    print_and_log("COMPLETE PENALTY ANALYSIS")
    print_and_log("="*60)
    
    total_penalty_impact = 0
    
    # 1. Health/Religion Violations
    print_and_log("\n1. HEALTH/RELIGION VIOLATIONS (-2048 points each):")
    print_and_log("-" * 50)
    health_violations = []
    for var in model.getVars():
        if var.varName.startswith('Slack_') and var.x > 0.5:
            health_violations.append(var.varName)
            total_penalty_impact -= 2048
            print_and_log(f"   {var.varName}: {var.x}")
    
    if not health_violations:
        print_and_log("   No health/religion violations")
    print_and_log(f"   Subtotal: {len(health_violations)} violations = {len(health_violations) * -2048} points")
    
    # 2. Format Preference Violations
    print_and_log("\n2. FORMAT PREFERENCE VIOLATIONS (-8 points each):")
    print_and_log("-" * 50)
    format_violations = []
    for var in model.getVars():
        if var.varName.startswith('Penalty_') and not var.varName.startswith('ConsecutivePenalty_') and var.x > 0.5:
            format_violations.append(var.varName)
            total_penalty_impact -= 8
            print_and_log(f"   {var.varName}: {var.x}")
    
    if not format_violations:
        print_and_log("   No format preference violations")
    print_and_log(f"   Subtotal: {len(format_violations)} violations = {len(format_violations) * -8} points")
    
    # 3. Consecutive Slot Aversion
    print_and_log("\n3. CONSECUTIVE SLOT AVERSION VIOLATIONS (-2048 points each):")
    print_and_log("-" * 50)
    consecutive_violations = []
    for var in model.getVars():
        if var.varName.startswith('ConsecutivePenalty_') and var.x > 0.5:
            consecutive_violations.append(var.varName)
            total_penalty_impact -= 2048
            print_and_log(f"   {var.varName}: {var.x}")
    
    if not consecutive_violations:
        print_and_log("   No consecutive slot aversion violations")
    print_and_log(f"   Subtotal: {len(consecutive_violations)} violations = {len(consecutive_violations) * -2048} points")
    
    # 4. Day Condensation (Z_ variables)
    print_and_log("\n4. DAY CONDENSATION IMPACT:")
    print_and_log("-" * 50)
    
    try:
        general_preferences_df = pd.read_excel(excel_file, sheet_name="General Preferences")
        if len(general_preferences_df.columns) > 3:
            general_preferences_df = general_preferences_df.rename(columns={
                general_preferences_df.columns[1]: 'Email',
                general_preferences_df.columns[3]: 'Day Preference'
            })
            day_preference_dict = general_preferences_df.set_index('Email')['Day Preference'].to_dict()
        else:
            day_preference_dict = {}
    except:
        day_preference_dict = {}
    
    day_impact = 0
    for var in model.getVars():
        if var.varName.startswith('Z_') and var.x > 0.5:
            parts = var.varName.split('_')
            if len(parts) >= 3:
                instructor = parts[1]
                day = parts[2]
                
                instructor_rows = df[df['Instructor'] == instructor]
                if not instructor_rows.empty:
                    email = instructor_rows['Email'].iloc[0]
                    prefers_condensed = day_preference_dict.get(email, "No") == "I prefer to condense my sections into fewer days"
                    penalty = -8 if prefers_condensed else -3
                    day_impact += penalty
                    total_penalty_impact += penalty
                    
                    pref_str = "prefers condensed" if prefers_condensed else "doesn't prefer condensed"
                    print_and_log(f"   {instructor} teaching on {day} ({pref_str}): {penalty} points")
    
    print_and_log(f"   Subtotal: {day_impact} points")
    
    # 5. Individual Preference Scheduling Impact
    print_and_log("\n5. INDIVIDUAL PREFERENCE SCHEDULING IMPACT:")
    print_and_log("-" * 50)
    
    preference_impact = 0
    try:
        df_constraints = pd.read_excel(excel_file, sheet_name='Constraints & Preferences')
        
        time_slot_mapping = {'M': 'Monday', 'T': 'Tuesday', 'W': 'Wednesday', 'R': 'Thursday', 'F': 'Friday'}
        time_slot_index = {
            '1': "8:30-10:00 AM", '2': "10:00-11:30 AM", '3': "11:30-1:00 PM", '4': "1:00-2:30 PM",
            '5': "2:30-4:00 PM", '6': "4:00-5:30 PM", '7': "6:00-7:30 PM", '8': "7:30-9:00 PM"
        }
        
        instructor_impacts = {}
        
        for _, row in df_constraints.iterrows():
            instructor_info = row['Instructor UCID: Type']
            slots = row['Slots']
            if isinstance(instructor_info, float):
                break
            
            email, constraint_type = instructor_info.split(": ")
            
            if constraint_type.strip() == "Pref-1":
                points = 8
            elif constraint_type.strip() == "Pref-2":
                points = 4
            elif constraint_type.strip() == "Pref-3":
                points = 2
            elif constraint_type.strip() == "Childcare":
                points = -1024
            elif constraint_type.strip() in ["Health", "Religion"]:
                continue
            else:
                points = -8
            
            instructor_row = df[df['Email'] == email]
            if instructor_row.empty:
                continue
            instructor_name = instructor_row['Instructor'].iloc[0]
            
            if instructor_name not in instructor_impacts:
                instructor_impacts[instructor_name] = 0
            
            blocked_slots = slots.split("|")[1:-1]
            for slot_code in blocked_slots:
                if len(slot_code) < 2:
                    continue
                
                day_abbrev = slot_code[0]
                time_slot_num = slot_code[1]
                
                day_full = time_slot_mapping[day_abbrev]
                time_slot_full = time_slot_index[time_slot_num]
                
                for var in model.getVars():
                    if (var.varName.startswith('X_') and var.x > 0.5 and 
                        instructor_name in var.varName and 
                        day_full in var.varName and 
                        time_slot_full in var.varName):
                        
                        instructor_impacts[instructor_name] += points
                        preference_impact += points
                        
                        impact_type = "POSITIVE" if points > 0 else "NEGATIVE"
                        print_and_log(f"   {instructor_name} scheduled in {constraint_type} slot {day_full} {time_slot_full}: {points} points ({impact_type})")
        
        if instructor_impacts:
            print_and_log("\n   INSTRUCTOR PREFERENCE SUMMARY:")
            for instructor, impact in sorted(instructor_impacts.items(), key=lambda x: x[1]):
                print_and_log(f"   {instructor}: {impact} points")
        
        total_penalty_impact += preference_impact
        print_and_log(f"   Subtotal: {preference_impact} points")
        
    except Exception as e:
        print_and_log(f"   Could not analyze individual preferences: {e}")
    
    print_and_log("\n" + "="*60)
    print_and_log(f"TOTAL PENALTY IMPACT: {total_penalty_impact} points")
    print_and_log(f"MODEL OBJECTIVE VALUE: {model.ObjVal:.0f}")
    print_and_log("="*60)
    
    filename = f"complete_penalty_analysis_{date_time_str}.txt"
    with open(filename, "w") as f:
        for line in output_lines:
            f.write(line + "\n")
    
    print(f"\n📄 Complete penalty analysis saved to: {filename}")
    return filename

if __name__ == "__main__":
    main()
