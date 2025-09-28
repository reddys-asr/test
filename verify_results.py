"""
Verification script to check the meeting scheduling results
"""

import pandas as pd

def verify_results():
    """Verify the meeting scheduling results"""
    
    output_file = r"d:\babai\Consolidated_Scheduled_Final.xlsx"
    
    try:
        # Load the output file
        associate_roster = pd.read_excel(output_file, sheet_name='Associate_Roster')
        manager_roster = pd.read_excel(output_file, sheet_name='Manager_Roster')
        
        print("="*60)
        print("MEETING SCHEDULING VERIFICATION")
        print("="*60)
        
        # Check associate scheduling
        working_associates = associate_roster[associate_roster['Working'] == 1]
        
        team_huddle_scheduled = working_associates[working_associates['Team_Huddle'].notna() & (working_associates['Team_Huddle'] != '')].shape[0]
        one_to_one_scheduled = working_associates[working_associates['One-2-One'].notna() & (working_associates['One-2-One'] != '')].shape[0]
        
        print(f"\n📊 ASSOCIATE SCHEDULING RESULTS:")
        print(f"   Total working associates: {len(working_associates)}")
        print(f"   Team Huddle scheduled: {team_huddle_scheduled}")
        print(f"   One-2-One scheduled: {one_to_one_scheduled}")
        
        # Show sample of scheduled associates
        print(f"\n👥 SAMPLE SCHEDULED ASSOCIATES:")
        sample_associates = working_associates[['AA_Name', 'Date', 'Manager', 'Team_Huddle', 'One-2-One']].head(10)
        print(sample_associates.to_string(index=False))
        
        # Check associates without One-2-One
        missing_one_to_one = working_associates[
            working_associates['One-2-One'].isna() | (working_associates['One-2-One'] == '')
        ]
        
        print(f"\n❌ ASSOCIATES WITHOUT ONE-2-ONE ({len(missing_one_to_one)}):")
        if len(missing_one_to_one) > 0:
            print(missing_one_to_one[['AA_Name', 'Date', 'Manager', 'Working', 'Shift_start']].head(10).to_string(index=False))
        else:
            print("   All working associates have One-2-One meetings scheduled!")
        
        # Check manager scheduling
        working_managers = manager_roster[manager_roster['Working'] == 1]
        
        # Check if Team_Huddle column exists
        if 'Team_Huddle' in manager_roster.columns:
            manager_team_huddle = working_managers[working_managers['Team_Huddle'].notna() & (working_managers['Team_Huddle'] != '')].shape[0]
        else:
            manager_team_huddle = 0
            
        manager_one_to_one = working_managers[working_managers['One-2-One'].notna() & (working_managers['One-2-One'] != '')].shape[0]
        
        print(f"\n👔 MANAGER SCHEDULING RESULTS:")
        print(f"   Total working managers: {len(working_managers)}")
        print(f"   Managers with Team Huddle: {manager_team_huddle}")
        print(f"   Managers with One-2-One: {manager_one_to_one}")
        
        # Show sample of scheduled managers
        print(f"\n👔 SAMPLE SCHEDULED MANAGERS:")
        if 'Team_Huddle' in manager_roster.columns:
            sample_managers = working_managers[['Manager', 'Date', 'Team_Huddle', 'One-2-One']].head(10)
        else:
            sample_managers = working_managers[['Manager', 'Date', 'One-2-One']].head(10)
        print(sample_managers.to_string(index=False))
        
        # Check for data consistency
        print(f"\n🔍 DATA CONSISTENCY CHECKS:")
        
        # Check if manager meetings match associate meetings
        associate_one_to_one_count = len(working_associates[working_associates['One-2-One'].notna() & (working_associates['One-2-One'] != '')])
        manager_one_to_one_entries = working_managers['One-2-One'].notna() & (working_managers['One-2-One'] != '')
        manager_meeting_count = working_managers[manager_one_to_one_entries]['One-2-One'].count()
        
        print(f"   Associates with One-2-One: {associate_one_to_one_count}")
        print(f"   Manager One-2-One entries: {manager_meeting_count}")
        
        # Check unique dates and managers
        unique_dates = associate_roster['Date'].nunique()
        unique_managers = manager_roster['Manager'].nunique()
        
        print(f"   Unique dates in data: {unique_dates}")
        print(f"   Unique managers: {unique_managers}")
        
        print("\n" + "="*60)
        
        return True
        
    except Exception as e:
        print(f"❌ Error verifying results: {str(e)}")
        return False

if __name__ == "__main__":
    verify_results()