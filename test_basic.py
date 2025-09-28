"""
Simple test to validate data loading and basic structure
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import openpyxl

def test_data_loading():
    """Test basic data loading functionality"""
    
    input_file = r"d:\babai\Consolidated.xlsx"
    
    print("🔍 Testing data loading...")
    
    try:
        # Load all sheets
        constraint_data = pd.read_excel(input_file, sheet_name='Constraint')
        associate_roster = pd.read_excel(input_file, sheet_name='Associate_Roster')
        manager_roster = pd.read_excel(input_file, sheet_name='Manager_Roster')
        schedule_heatmap = pd.read_excel(input_file, sheet_name='Schedule_heatmap')
        
        print("✅ All sheets loaded successfully")
        print(f"   - Constraint: {len(constraint_data)} rows")
        print(f"   - Associates: {len(associate_roster)} rows")
        print(f"   - Managers: {len(manager_roster)} rows")
        print(f"   - Heatmap: {len(schedule_heatmap)} rows")
        
        # Get NPT threshold
        wb = openpyxl.load_workbook(input_file)
        constraint_sheet = wb['Constraint']
        npt_threshold = constraint_sheet['G1'].value
        wb.close()
        
        print(f"   - NPT Threshold from G1: {npt_threshold}")
        
        # Show constraint data structure
        print("\n📋 Constraint Data:")
        print(constraint_data.to_string())
        
        # Show sample associate data
        print("\n👥 Sample Associate Data:")
        print(associate_roster[['Date', 'AA_Name', 'Manager', 'Shift_start', 'Working']].head())
        
        # Filter working associates
        working_associates = associate_roster[associate_roster['Working'] == 1]
        print(f"\n💼 Working Associates: {len(working_associates)} out of {len(associate_roster)}")
        
        # Show unique shift start times
        unique_shifts = working_associates['Shift_start'].unique()
        print(f"📅 Unique Shift Start Times: {len(unique_shifts)}")
        for shift in unique_shifts[:5]:  # Show first 5
            print(f"   - {shift}")
        
        # Test basic scheduling logic
        print("\n🧪 Testing basic scheduling logic...")
        
        # Get Team Huddle meetings
        team_huddle_meetings = constraint_data[
            constraint_data['Meeting_Name'].str.contains('Team_Huddle', case=False, na=False)
        ]
        
        print(f"🤝 Team Huddle meetings found: {len(team_huddle_meetings)}")
        if len(team_huddle_meetings) > 0:
            print(team_huddle_meetings.to_string())
        
        # Get One-to-One meetings
        one_to_one_meetings = constraint_data[
            ~constraint_data['Meeting_Name'].str.contains('Team_Huddle', case=False, na=False)
        ]
        
        print(f"👥 One-to-One meetings found: {len(one_to_one_meetings)}")
        if len(one_to_one_meetings) > 0:
            print(one_to_one_meetings.to_string())
        
        return True
        
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        return False

def test_time_operations():
    """Test time-related operations"""
    
    print("\n⏰ Testing time operations...")
    
    try:
        # Test time parsing
        sample_time = "06:30:00"
        parsed_time = pd.to_datetime(sample_time, format='%H:%M:%S').time()
        print(f"✅ Time parsing: {sample_time} -> {parsed_time}")
        
        # Test datetime operations
        shift_start = datetime.strptime("06:30:00", "%H:%M:%S")
        first_interval = shift_start
        second_interval = shift_start + timedelta(minutes=30)
        
        print(f"✅ Interval calculation:")
        print(f"   - Shift start: {first_interval.strftime('%H:%M')}")
        print(f"   - Second interval: {second_interval.strftime('%H:%M')}")
        
        return True
        
    except Exception as e:
        print(f"❌ Time operation error: {str(e)}")
        return False

if __name__ == "__main__":
    print("="*60)
    print("BASIC DATA LOADING TEST")
    print("="*60)
    
    success1 = test_data_loading()
    success2 = test_time_operations()
    
    if success1 and success2:
        print("\n🎉 Basic tests passed! Data structure is valid.")
    else:
        print("\n💥 Some basic tests failed.")
    
    print("\n" + "="*60)