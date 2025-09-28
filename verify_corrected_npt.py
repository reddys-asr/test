"""
NPT Count Verification - Corrected Durations

Verify NPT Count calculations with correct meeting durations:
- Team Huddle: 15 minutes
- One-2-One: 30 minutes

Formula: NPT Count = (meetings × duration) / 30
"""

import pandas as pd

def verify_corrected_npt_calculation():
    """Verify NPT Count calculations with correct durations"""
    
    # Read the results
    schedule_heatmap = pd.read_excel('Consolidated_Scheduled_Final.xlsx', sheet_name='Schedule_heatmap')
    associate_roster = pd.read_excel('Consolidated_Scheduled_Final.xlsx', sheet_name='Associate_Roster')
    
    print("=" * 70)
    print("NPT COUNT VERIFICATION - CORRECTED DURATIONS")
    print("=" * 70)
    
    print(f"\n📋 MEETING DURATIONS:")
    print(f"   Team Huddle: 15 minutes → NPT Count = 15/30 = 0.5 per meeting")
    print(f"   One-2-One: 30 minutes → NPT Count = 30/30 = 1.0 per meeting")
    
    # Show sample NPT Count calculations
    print(f"\n📊 SAMPLE NPT COUNT CALCULATIONS:")
    sample_data = schedule_heatmap[schedule_heatmap['NPT Count'] > 0].head(10)
    print(sample_data[['Date', 'Skill', 'Interval', 'NPT Count', 'Scheduled', 'Revised Staffing']].to_string(index=False))
    
    # Manual verification for Team Huddle meetings
    print(f"\n🔍 MANUAL VERIFICATION - TEAM HUDDLE:")
    print(f"   Checking Skill 1 on 2025-09-14 at 01:30")
    
    skill1_team_huddles_0130 = associate_roster[
        (associate_roster['Skill'] == 'Skill 1') & 
        (associate_roster['Date'] == '2025-09-14') &
        (associate_roster['Team_Huddle'] == '01:30')
    ]
    
    expected_npt_team_huddle = len(skill1_team_huddles_0130) * 15 / 30  # 15 minutes each
    print(f"   Team Huddle meetings: {len(skill1_team_huddles_0130)}")
    print(f"   Expected NPT Count: {len(skill1_team_huddles_0130)} × 15 ÷ 30 = {expected_npt_team_huddle}")
    
    # Check for One-2-One meetings at the same time
    skill1_one_to_one_0130 = associate_roster[
        (associate_roster['Skill'] == 'Skill 1') & 
        (associate_roster['Date'] == '2025-09-14') &
        (associate_roster['One-2-One'] == '01:30')
    ]
    
    expected_npt_one_to_one = len(skill1_one_to_one_0130) * 30 / 30  # 30 minutes each
    print(f"   One-2-One meetings: {len(skill1_one_to_one_0130)}")
    print(f"   Expected NPT Count: {len(skill1_one_to_one_0130)} × 30 ÷ 30 = {expected_npt_one_to_one}")
    
    total_expected_npt = expected_npt_team_huddle + expected_npt_one_to_one
    print(f"   Total Expected NPT Count: {expected_npt_team_huddle} + {expected_npt_one_to_one} = {total_expected_npt}")
    
    # Check actual NPT Count in heatmap
    heatmap_entry = schedule_heatmap[
        (schedule_heatmap['Skill'] == 'Skill 1') & 
        (schedule_heatmap['Date'] == '2025-09-14') &
        (schedule_heatmap['Interval'] == '01:30:00')
    ]
    
    if not heatmap_entry.empty:
        actual_npt = heatmap_entry['NPT Count'].iloc[0]
        print(f"   Actual NPT Count in heatmap: {actual_npt}")
        print(f"   ✅ Calculation {'CORRECT' if abs(actual_npt - total_expected_npt) < 0.001 else 'INCORRECT'}")
    else:
        print(f"   ❌ No heatmap entry found")
    
    # Check a few more examples with different combinations
    print(f"\n🔍 ADDITIONAL VERIFICATION EXAMPLES:")
    
    # Find entries with both Team Huddle and One-2-One
    mixed_intervals = []
    for _, row in schedule_heatmap[schedule_heatmap['NPT Count'] > 0].head(20).iterrows():
        date = row['Date']
        skill = row['Skill']
        interval = str(row['Interval']).replace(':00', '')  # Convert 01:30:00 to 01:30
        
        team_count = len(associate_roster[
            (associate_roster['Skill'] == skill) & 
            (associate_roster['Date'] == date) &
            (associate_roster['Team_Huddle'] == interval)
        ])
        
        one_to_one_count = len(associate_roster[
            (associate_roster['Skill'] == skill) & 
            (associate_roster['Date'] == date) &
            (associate_roster['One-2-One'] == interval)
        ])
        
        expected = (team_count * 15 + one_to_one_count * 30) / 30
        actual = row['NPT Count']
        
        print(f"   {date.strftime('%Y-%m-%d')} {skill} {interval}: {team_count}TH + {one_to_one_count}O2O = {expected:.1f} (actual: {actual:.1f})")
        
        if len(mixed_intervals) >= 5:  # Show only first 5 examples
            break
    
    # Summary statistics
    print(f"\n📈 NPT COUNT STATISTICS:")
    print(f"   Total heatmap entries: {len(schedule_heatmap)}")
    print(f"   Entries with NPT Count > 0: {len(schedule_heatmap[schedule_heatmap['NPT Count'] > 0])}")
    print(f"   Max NPT Count: {schedule_heatmap['NPT Count'].max():.2f}")
    print(f"   Average NPT Count: {schedule_heatmap['NPT Count'].mean():.3f}")
    print(f"   Min NPT Count (non-zero): {schedule_heatmap[schedule_heatmap['NPT Count'] > 0]['NPT Count'].min():.2f}")
    
    # Show decimal examples
    decimal_npt = schedule_heatmap[(schedule_heatmap['NPT Count'] % 1 != 0) & (schedule_heatmap['NPT Count'] > 0)]
    print(f"\n🔢 DECIMAL NPT COUNT EXAMPLES (Team Huddles):")
    if len(decimal_npt) > 0:
        print(decimal_npt[['Date', 'Skill', 'Interval', 'NPT Count']].head().to_string(index=False))
    else:
        print("   No decimal NPT counts found")
    
    print("=" * 70)

if __name__ == "__main__":
    verify_corrected_npt_calculation()