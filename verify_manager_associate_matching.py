"""
Detailed Manager-Associate Matching Verification

This script specifically checks that One-2-One meetings are scheduled 
with the correct manager-associate pairs on their shared working days.
"""

import pandas as pd
import numpy as np
from datetime import datetime

def verify_manager_associate_matching():
    """Verify One-2-One meetings match correct manager-associate pairs"""
    
    # Read the scheduled results
    associate_roster = pd.read_excel('Consolidated_Scheduled_Final.xlsx', sheet_name='Associate_Roster')
    manager_roster = pd.read_excel('Consolidated_Scheduled_Final.xlsx', sheet_name='Manager_Roster')
    
    print("=" * 60)
    print("MANAGER-ASSOCIATE MATCHING VERIFICATION")
    print("=" * 60)
    
    # Filter for associates with One-2-One meetings
    associates_with_meetings = associate_roster[
        (associate_roster['Working'] == 1) & 
        (pd.notna(associate_roster['One-2-One']))
    ].copy()
    
    print(f"\n📊 SCHEDULED ONE-2-ONE MEETINGS: {len(associates_with_meetings)}")
    
    # Check each One-2-One meeting
    correct_matches = 0
    total_matches = 0
    verification_details = []
    
    for idx, associate in associates_with_meetings.iterrows():
        total_matches += 1
        
        associate_name = associate['AA_Name']
        assigned_manager = associate['Manager']
        meeting_date = associate['Date']
        meeting_time = associate['One-2-One']
        
        # Find the corresponding manager on the same date
        manager_on_date = manager_roster[
            (manager_roster['Manager'] == assigned_manager) & 
            (manager_roster['Date'] == meeting_date) &
            (manager_roster['Working'] == 1)
        ]
        
        if not manager_on_date.empty:
            # Check if manager has this meeting time in their schedule
            manager_meetings = manager_on_date['One-2-One'].iloc[0]
            
            if pd.notna(manager_meetings):
                # Parse manager's meeting times (could be comma-separated)
                manager_times = str(manager_meetings).split(', ')
                manager_times = [t.strip() for t in manager_times if t.strip()]
                
                if meeting_time in manager_times:
                    correct_matches += 1
                    match_status = "✅ CORRECT"
                else:
                    match_status = "❌ TIME_MISMATCH"
            else:
                match_status = "❌ MANAGER_NO_MEETING"
        else:
            match_status = "❌ MANAGER_NOT_WORKING"
        
        verification_details.append({
            'Associate': associate_name,
            'Manager': assigned_manager,
            'Date': meeting_date,
            'Meeting_Time': meeting_time,
            'Status': match_status
        })
    
    # Summary
    print(f"\n🎯 MATCHING RESULTS:")
    print(f"   Correct matches: {correct_matches}")
    print(f"   Total matches: {total_matches}")
    print(f"   Success rate: {correct_matches/total_matches*100:.1f}%")
    
    # Show sample verification details
    print(f"\n📋 SAMPLE MATCHING DETAILS:")
    verification_df = pd.DataFrame(verification_details)
    print(verification_df.head(15).to_string(index=False))
    
    # Show any mismatches
    mismatches = verification_df[~verification_df['Status'].str.contains('CORRECT')]
    if not mismatches.empty:
        print(f"\n❌ MISMATCHES FOUND ({len(mismatches)}):")
        print(mismatches.to_string(index=False))
    else:
        print(f"\n✅ ALL MEETINGS PROPERLY MATCHED!")
    
    # Check manager availability distribution
    print(f"\n👔 MANAGER UTILIZATION:")
    manager_counts = associates_with_meetings['Manager'].value_counts()
    print(f"   Managers with One-2-One meetings: {len(manager_counts)}")
    print(f"   Average meetings per manager: {manager_counts.mean():.1f}")
    print(f"   Top 5 busy managers:")
    for manager, count in manager_counts.head().items():
        print(f"      {manager}: {count} meetings")
    
    # Check date distribution
    print(f"\n📅 DATE DISTRIBUTION:")
    date_counts = associates_with_meetings['Date'].value_counts().sort_index()
    for date, count in date_counts.items():
        print(f"   {date}: {count} meetings")
    
    print("=" * 60)

if __name__ == "__main__":
    verify_manager_associate_matching()