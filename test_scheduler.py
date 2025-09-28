"""
Test script to validate the meeting scheduler functionality
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os

def test_meeting_scheduler():
    """Test the meeting scheduler with basic validation"""
    
    try:
        from meeting_scheduler import MeetingScheduler
        print("✅ Meeting scheduler module imported successfully")
        
        # Test file paths
        input_file = r"d:\babai\Consolidated.xlsx"
        output_file = r"d:\babai\Consolidated_Scheduled.xlsx"
        
        # Check if input file exists
        if not os.path.exists(input_file):
            print(f"❌ Input file not found: {input_file}")
            print("📝 Please ensure the Consolidated.xlsx file exists at the specified location")
            return False
        
        print(f"✅ Input file found: {input_file}")
        
        # Initialize scheduler
        scheduler = MeetingScheduler(input_file, output_file)
        print("✅ Meeting scheduler initialized")
        
        # Test data loading
        try:
            scheduler.load_data()
            print("✅ Data loaded successfully")
            print(f"   - NPT Threshold: {scheduler.npt_threshold}")
            print(f"   - Input Data: {len(scheduler.input_data)} rows")
            print(f"   - Associates: {len(scheduler.associate_roster)} rows")
            print(f"   - Managers: {len(scheduler.manager_roster)} rows")
            print(f"   - Heatmap: {len(scheduler.schedule_heatmap)} rows")
            
        except Exception as e:
            print(f"❌ Error loading data: {str(e)}")
            return False
        
        # Test scheduling methods individually
        try:
            print("\n🔄 Testing Team Huddle scheduling...")
            scheduler.schedule_team_huddles()
            print("✅ Team Huddle scheduling completed")
            
            print("\n🔄 Testing One-to-One meeting scheduling...")
            scheduler.schedule_one_to_one_meetings()
            print("✅ One-to-One meeting scheduling completed")
            
            print("\n🔄 Testing Manager Roster updates...")
            scheduler.update_manager_roster_meetings()
            print("✅ Manager Roster updated")
            
            print("\n🔄 Testing Schedule Heatmap updates...")
            scheduler.update_schedule_heatmap()
            print("✅ Schedule Heatmap updated")
            
            print("\n🔄 Generating summary report...")
            summary = scheduler.generate_summary_report()
            print("✅ Summary report generated")
            
            print("\n🔄 Saving results...")
            scheduler.save_results()
            print("✅ Results saved successfully")
            
            # Print basic summary
            print(f"\n📊 QUICK SUMMARY:")
            print(f"   Total meetings: {summary['total_meetings_scheduled']}")
            print(f"   Unscheduled: {summary['unscheduled_meetings']}")
            
            return True
            
        except Exception as e:
            print(f"❌ Error in scheduling process: {str(e)}")
            return False
        
    except ImportError as e:
        print(f"❌ Import error: {str(e)}")
        return False
    
    except Exception as e:
        print(f"❌ Unexpected error: {str(e)}")
        return False

def validate_environment():
    """Validate the Python environment and required packages"""
    print("🔍 Validating environment...")
    
    required_packages = ['pandas', 'numpy', 'openpyxl']
    
    for package in required_packages:
        try:
            __import__(package)
            print(f"✅ {package} available")
        except ImportError:
            print(f"❌ {package} not available")
            return False
    
    return True

if __name__ == "__main__":
    print("="*60)
    print("MEETING SCHEDULER TEST SUITE")
    print("="*60)
    
    # Validate environment
    if not validate_environment():
        print("\n❌ Environment validation failed")
        exit(1)
    
    print("\n🧪 Running meeting scheduler test...")
    
    # Run test
    success = test_meeting_scheduler()
    
    if success:
        print("\n🎉 All tests passed! Meeting scheduler is ready to use.")
    else:
        print("\n💥 Some tests failed. Please check the error messages above.")
    
    print("\n" + "="*60)