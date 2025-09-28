"""
NPT Count Verification Script

Verify that the NPT Count calculations are correct according to the formula:
NPT Count = (All meetings scheduled in that interval based on day level and skill level × MeetingDuration / 30)
"""

import pandas as pd

def verify_npt_count_calculation():
    """Verify NPT Count calculations in the updated heatmap"""
    
    # Read the results
    schedule_heatmap = pd.read_excel('Consolidated_Scheduled_Final.xlsx', sheet_name='Schedule_heatmap')
    associate_roster = pd.read_excel('Consolidated_Scheduled_Final.xlsx', sheet_name='Associate_Roster')
    
    print("=" * 60)
    print("NPT COUNT CALCULATION VERIFICATION")
    print("=" * 60)
    
    # Show sample NPT Count calculations
    print(f"\n📊 SAMPLE NPT COUNT CALCULATIONS:")
    sample_data = schedule_heatmap[schedule_heatmap['NPT Count'] > 0].head(10)
    print(sample_data[['Date', 'Skill', 'Interval', 'NPT Count', 'Scheduled', 'Revised Staffing']].to_string(index=False))
    
    # Verify a specific calculation manually
    print(f"\n🔍 MANUAL VERIFICATION FOR SKILL 1 ON 2025-09-14 AT 01:30:")
    
    # Count Team Huddle meetings for Skill 1 on 2025-09-14 at 01:30
    skill1_associates = associate_roster[
        (associate_roster['Skill'] == 'Skill 1') & 
        (associate_roster['Date'] == '2025-09-14') &
        (associate_roster['Team_Huddle'] == '01:30')
    ]
    
    print(f"   Associates with Team Huddle at 01:30: {len(skill1_associates)}")
    print(f"   Team Huddle duration: 30 minutes")
    print(f"   Expected NPT Count: {len(skill1_associates)} × 30 / 30 = {len(skill1_associates)}")
    
    # Check actual NPT Count in heatmap
    heatmap_entry = schedule_heatmap[
        (schedule_heatmap['Skill'] == 'Skill 1') & 
        (schedule_heatmap['Date'] == '2025-09-14') &
        (schedule_heatmap['Interval'] == '01:30:00')
    ]
    
    if not heatmap_entry.empty:
        actual_npt = heatmap_entry['NPT Count'].iloc[0]
        print(f"   Actual NPT Count in heatmap: {actual_npt}")
        print(f"   ✅ Calculation {'CORRECT' if actual_npt == len(skill1_associates) else 'INCORRECT'}")
    else:
        print(f"   ❌ No heatmap entry found for this combination")
    
    # Summary statistics
    print(f"\n📈 NPT COUNT STATISTICS:")
    print(f"   Total heatmap entries: {len(schedule_heatmap)}")
    print(f"   Entries with NPT Count > 0: {len(schedule_heatmap[schedule_heatmap['NPT Count'] > 0])}")
    print(f"   Max NPT Count: {schedule_heatmap['NPT Count'].max()}")
    print(f"   Average NPT Count: {schedule_heatmap['NPT Count'].mean():.2f}")
    
    # Check for negative Revised Staffing
    negative_staffing = schedule_heatmap[schedule_heatmap['Revised Staffing'] < 0]
    print(f"\n⚠️  REVISED STAFFING WARNINGS:")
    print(f"   Entries with negative Revised Staffing: {len(negative_staffing)}")
    print(f"   Min Revised Staffing: {schedule_heatmap['Revised Staffing'].min():.2f}")
    
    if len(negative_staffing) > 0:
        print(f"   Sample negative staffing entries:")
        print(negative_staffing[['Date', 'Skill', 'Interval', 'NPT Count', 'Scheduled', 'Revised Staffing']].head().to_string(index=False))
    
    print("=" * 60)

if __name__ == "__main__":
    verify_npt_count_calculation()