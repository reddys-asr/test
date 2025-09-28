"""
Meeting Scheduler - Comprehensive System for Associate and Manager Meeting Scheduling

This script reads from Consolidated.xlsx and schedules meetings based on:
1. Team Huddle scheduling within first hour of shift (50-60% distribution)
2. One-to-One meetings with manager availability and frequency constraints
3. Heatmap constraints for NPT count and revised staffing validation
4. Output generation with summary reports

Author: GitHub Copilot
Date: September 28, 2025
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import openpyxl
from openpyxl import Workbook
import logging
from collections import defaultdict
import random
import copy

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class MeetingScheduler:
    def __init__(self, input_file_path, output_file_path):
        """
        Initialize the Meeting Scheduler
        
        Args:
            input_file_path (str): Path to input Consolidated.xlsx file
            output_file_path (str): Path to output file with scheduled meetings
        """
        self.input_file_path = input_file_path
        self.output_file_path = output_file_path
        
        # Data containers
        self.input_data = None
        self.associate_roster = None
        self.manager_roster = None
        self.schedule_heatmap = None
        
        # Configuration from Input sheet
        self.npt_threshold = None
        self.meeting_types = []
        
        # Tracking containers
        self.scheduled_meetings = defaultdict(list)  # associate_id -> [meetings]
        self.manager_meetings = defaultdict(list)    # manager_id -> [meetings]
        self.unscheduled_meetings = []
        self.team_huddle_stats = defaultdict(int)
        
        logger.info(f"Meeting Scheduler initialized with input: {input_file_path}")
    
    def load_data(self):
        """Load all sheets from the Excel file and validate structure"""
        try:
            logger.info("Loading data from Excel file...")
            
            # Load all sheets
            self.input_data = pd.read_excel(self.input_file_path, sheet_name='Constraint')
            self.associate_roster = pd.read_excel(self.input_file_path, sheet_name='Associate_Roster')
            self.manager_roster = pd.read_excel(self.input_file_path, sheet_name='Manager_Roster')
            self.schedule_heatmap = pd.read_excel(self.input_file_path, sheet_name='Schedule_heatmap')
            
            # Extract NPT threshold from Constraint sheet (G1 cell)
            wb = openpyxl.load_workbook(self.input_file_path)
            input_sheet = wb['Constraint']
            self.npt_threshold = input_sheet['G1'].value or 5
            wb.close()
            
            logger.info(f"NPT Threshold: {self.npt_threshold}")
            logger.info(f"Loaded {len(self.input_data)} meeting types")
            logger.info(f"Loaded {len(self.associate_roster)} associates")
            logger.info(f"Loaded {len(self.manager_roster)} managers")
            logger.info(f"Loaded {len(self.schedule_heatmap)} heatmap entries")
            
            self._validate_data_structure()
            self._prepare_data()
            
        except Exception as e:
            logger.error(f"Error loading data: {str(e)}")
            raise
    
    def _validate_data_structure(self):
        """Validate the structure of loaded data"""
        logger.info("Validating data structure...")
        
        # Validate Constraint sheet columns
        required_constraint_cols = ['Meeting_Name', 'Frequency', 'Meeting_Type', 
                                  'Manager_Availability', 'Duration', 'NPT_Threshold']
        
        # Validate Associate_Roster columns
        required_associate_cols = ['Date', 'AA_Name', 'Day', 'Manager', 'site', 
                                 'Shift_start', 'lunch1_start', 'lunch1_end',
                                 'break1_start', 'break1_end', 'break2_start', 'break2_end', 'Working']
        
        # Check if required columns exist (flexible matching)
        for col in required_associate_cols:
            if col not in self.associate_roster.columns:
                logger.warning(f"Column '{col}' not found in Associate_Roster")
        
        logger.info("Data structure validation completed")
    
    def _prepare_data(self):
        """Prepare and clean data for processing"""
        logger.info("Preparing data for processing...")
        
        # Filter working associates only
        self.associate_roster = self.associate_roster[self.associate_roster['Working'] == 1].copy()
        
        # Filter working managers only
        self.manager_roster = self.manager_roster[self.manager_roster['Working'] == 1].copy()
        
        # Convert time columns to datetime objects
        time_columns = ['Shift_start', 'lunch1_start', 'lunch1_end', 
                       'break1_start', 'break1_end', 'break2_start', 'break2_end']
        
        for col in time_columns:
            if col in self.associate_roster.columns:
                self.associate_roster[col] = pd.to_datetime(self.associate_roster[col], errors='coerce')
            if col in self.manager_roster.columns:
                self.manager_roster[col] = pd.to_datetime(self.manager_roster[col], errors='coerce')
        
        # Use existing meeting columns if they exist, otherwise create new ones
        meeting_names = self.input_data['Meeting_Name'].tolist()
        for meeting_name in meeting_names:
            if meeting_name not in self.associate_roster.columns:
                self.associate_roster[meeting_name] = ''
            if meeting_name not in self.manager_roster.columns:
                self.manager_roster[meeting_name] = ''
        
        logger.info("Data preparation completed")
    
    def get_30min_intervals(self, start_time, duration_hours=8):
        """Generate 30-minute intervals for the shift"""
        intervals = []
        current = start_time
        end_time = start_time + timedelta(hours=duration_hours)
        
        while current < end_time:
            intervals.append(current)
            current += timedelta(minutes=30)
        
        return intervals
    
    def is_time_conflicting(self, associate_row, meeting_time, duration_minutes):
        """Check if meeting time conflicts with lunch or break periods"""
        meeting_end = meeting_time + timedelta(minutes=duration_minutes)
        
        # Check lunch1 conflict
        if pd.notna(associate_row['lunch1_start']) and pd.notna(associate_row['lunch1_end']):
            if not (meeting_end <= associate_row['lunch1_start'] or meeting_time >= associate_row['lunch1_end']):
                return True
        
        # Check break1 conflict
        if pd.notna(associate_row['break1_start']) and pd.notna(associate_row['break1_end']):
            if not (meeting_end <= associate_row['break1_start'] or meeting_time >= associate_row['break1_end']):
                return True
        
        # Check break2 conflict
        if pd.notna(associate_row['break2_start']) and pd.notna(associate_row['break2_end']):
            if not (meeting_end <= associate_row['break2_start'] or meeting_time >= associate_row['break2_end']):
                return True
        
        return False
    
    def schedule_team_huddles(self):
        """Schedule Team Huddle meetings within first hour of shift"""
        logger.info("Scheduling Team Huddle meetings...")
        
        team_huddle_meetings = self.input_data[
            self.input_data['Meeting_Name'].str.contains('Team_Huddle|Group', case=False, na=False)
        ]
        
        if team_huddle_meetings.empty:
            logger.info("No Team Huddle meetings found")
            return
        
        # Group associates by shift start time and site
        shift_groups = self.associate_roster.groupby(['Shift_start', 'site'])
        
        for (shift_start, site), group in shift_groups:
            associates = group[group['AA_Name'].str.startswith('AA', na=False)]
            
            if len(associates) == 0:
                continue
            
            logger.info(f"Scheduling Team Huddles for {len(associates)} associates starting at {shift_start}")
            
            # First hour intervals (30-min slots)
            first_interval = shift_start
            second_interval = shift_start + timedelta(minutes=30)
            
            # Calculate 50-60% distribution
            total_associates = len(associates)
            first_slot_count = int(total_associates * random.uniform(0.5, 0.6))
            second_slot_count = total_associates - first_slot_count
            
            # Randomly assign associates to slots
            associate_list = associates.index.tolist()
            random.shuffle(associate_list)
            
            first_slot_associates = associate_list[:first_slot_count]
            second_slot_associates = associate_list[first_slot_count:]
            
            # Schedule first interval
            for idx in first_slot_associates:
                for _, meeting in team_huddle_meetings.iterrows():
                    meeting_info = {
                        'associate_id': idx,
                        'meeting_type': meeting['Meeting_Name'],
                        'meeting_time': first_interval,
                        'duration': meeting['Duration'],
                        'site': site,
                        'date': associates.loc[idx, 'Date']
                    }
                    self.scheduled_meetings[idx].append(meeting_info)
                    
                    # Update associate roster with meeting time
                    self.associate_roster.loc[idx, meeting['Meeting_Name']] = first_interval.strftime('%H:%M')
            
            # Schedule second interval
            for idx in second_slot_associates:
                for _, meeting in team_huddle_meetings.iterrows():
                    meeting_info = {
                        'associate_id': idx,
                        'meeting_type': meeting['Meeting_Name'],
                        'meeting_time': second_interval,
                        'duration': meeting['Duration'],
                        'site': site,
                        'date': associates.loc[idx, 'Date']
                    }
                    self.scheduled_meetings[idx].append(meeting_info)
                    
                    # Update associate roster with meeting time
                    self.associate_roster.loc[idx, meeting['Meeting_Name']] = second_interval.strftime('%H:%M')
            
            # Track statistics
            self.team_huddle_stats[f"{shift_start}_{site}_first"] = first_slot_count
            self.team_huddle_stats[f"{shift_start}_{site}_second"] = second_slot_count
            self.team_huddle_stats[f"{shift_start}_{site}_total"] = total_associates
        
        logger.info("Team Huddle scheduling completed")
    
    def schedule_one_to_one_meetings(self):
        """Schedule One-to-One meetings with manager availability constraints"""
        logger.info("Scheduling One-to-One meetings...")
        
        one_to_one_meetings = self.input_data[
            ~self.input_data['Meeting Type'].str.contains('Team Huddle|Group', case=False, na=False)
        ]
        
        for _, meeting in one_to_one_meetings.iterrows():
            meeting_type = meeting['Meeting Type']
            frequency = meeting['Frequency']
            duration = meeting['Duration (minutes)']
            requires_direct_manager = meeting['Is Manager slots should available'] == 'Yes'
            
            logger.info(f"Scheduling {meeting_type} meetings ({frequency})")
            
            # Get associates who need this meeting
            for idx, associate in self.associate_roster.iterrows():
                if self._should_schedule_meeting(associate, meeting_type, frequency):
                    scheduled = self._schedule_individual_meeting(
                        idx, associate, meeting_type, duration, requires_direct_manager
                    )
                    
                    if scheduled:
                        # Update associate roster
                        col_index = list(self.input_data['Meeting Type']).index(meeting_type)
                        if col_index < 4:
                            col_name = ['P', 'Q', 'R', 'S'][col_index]
                            self.associate_roster.loc[idx, col_name] = scheduled['meeting_time'].strftime('%H:%M')
                    else:
                        self.unscheduled_meetings.append({
                            'associate_id': idx,
                            'meeting_type': meeting_type,
                            'reason': 'No available slot found'
                        })
        
        logger.info("One-to-One meeting scheduling completed")
    
    def _should_schedule_meeting(self, associate, meeting_type, frequency):
        """Determine if associate should have this meeting scheduled"""
        # For daily meetings, schedule for everyone
        if frequency.lower() == 'daily':
            return True
        
        # For weekly meetings, limit to 25% per day randomly
        elif frequency.lower() == 'weekly':
            return random.random() < 0.25
        
        # For monthly meetings, schedule randomly (can be refined based on requirements)
        elif frequency.lower() == 'monthly':
            return random.random() < 0.1  # 10% chance for monthly
        
        return False
    
    def _schedule_individual_meeting(self, associate_idx, associate, meeting_type, duration, requires_direct_manager):
        """Schedule individual meeting for associate"""
        shift_start = associate['shift_start_time']
        intervals = self.get_30min_intervals(shift_start)
        
        # Skip first hour for non-huddle meetings to avoid Team Huddle conflicts
        available_intervals = intervals[2:]  # Start from 1 hour after shift start
        
        for interval in available_intervals:
            # Check for conflicts with lunch/breaks
            if self.is_time_conflicting(associate, interval, duration):
                continue
            
            # Check for conflicts with existing meetings
            if self._has_meeting_conflict(associate_idx, interval, duration):
                continue
            
            # Check manager availability
            manager_available = self._check_manager_availability(
                associate, interval, duration, requires_direct_manager
            )
            
            if manager_available:
                # Schedule the meeting
                meeting_info = {
                    'associate_id': associate_idx,
                    'meeting_type': meeting_type,
                    'meeting_time': interval,
                    'duration': duration,
                    'workgroup': associate['Workgroup'],
                    'date': associate['Date'],
                    'manager_id': manager_available
                }
                
                self.scheduled_meetings[associate_idx].append(meeting_info)
                self.manager_meetings[manager_available].append(meeting_info)
                
                return meeting_info
        
        return None
    
    def _has_meeting_conflict(self, associate_idx, meeting_time, duration):
        """Check if associate already has a meeting at this time"""
        meeting_end = meeting_time + timedelta(minutes=duration)
        
        for existing_meeting in self.scheduled_meetings[associate_idx]:
            existing_start = existing_meeting['meeting_time']
            existing_end = existing_start + timedelta(minutes=existing_meeting['duration'])
            
            # Check for overlap
            if not (meeting_end <= existing_start or meeting_time >= existing_end):
                return True
        
        return False
    
    def _check_manager_availability(self, associate, meeting_time, duration, requires_direct_manager):
        """Check if a manager is available for the meeting"""
        if requires_direct_manager:
            # Find direct manager from TM column
            if 'TM' in associate and pd.notna(associate['TM']):
                manager_id = associate['TM']
                # Check if this manager is available
                if self._is_manager_available(manager_id, meeting_time, duration):
                    return manager_id
            return None
        else:
            # Any available manager can take this meeting
            for idx, manager in self.manager_roster.iterrows():
                if self._is_manager_available(idx, meeting_time, duration):
                    return idx
            return None
    
    def _is_manager_available(self, manager_id, meeting_time, duration):
        """Check if specific manager is available at the given time"""
        if manager_id not in self.manager_roster.index:
            return False
        
        manager = self.manager_roster.loc[manager_id]
        
        # Check conflicts with manager's lunch/breaks
        if self.is_time_conflicting(manager, meeting_time, duration):
            return False
        
        # Check conflicts with manager's existing meetings
        meeting_end = meeting_time + timedelta(minutes=duration)
        
        for existing_meeting in self.manager_meetings[manager_id]:
            existing_start = existing_meeting['meeting_time']
            existing_end = existing_start + timedelta(minutes=existing_meeting['duration'])
            
            if not (meeting_end <= existing_start or meeting_time >= existing_end):
                return False
        
        return True
    
    def update_schedule_heatmap(self):
        """Update Schedule_heatmap with NPT Count and Revised Staffing"""
        logger.info("Updating Schedule_heatmap with NPT calculations...")
        
        # Add new columns
        self.schedule_heatmap['NPT Count'] = 0.0
        self.schedule_heatmap['Revised Staffing'] = 0.0
        
        # Calculate NPT Count for each interval
        for idx, row in self.schedule_heatmap.iterrows():
            date = row['Date']
            workgroup = row['Workgroup']
            interval_time = pd.to_datetime(row['Interval'])
            
            npt_count = 0
            
            # Count meetings in this interval
            for associate_meetings in self.scheduled_meetings.values():
                for meeting in associate_meetings:
                    if (meeting['date'] == date and 
                        meeting['workgroup'] == workgroup and
                        meeting['meeting_time'] == interval_time):
                        # NPT Count = (meetings × duration) / 30
                        npt_count += meeting['duration'] / 30
            
            self.schedule_heatmap.loc[idx, 'NPT Count'] = npt_count
            
            # Calculate Revised Staffing = (Scheduled - NPT Count) - Requirement
            scheduled = row['Scheduled']
            requirement = row['Requirement']
            revised_staffing = (scheduled - npt_count) - requirement
            
            self.schedule_heatmap.loc[idx, 'Revised Staffing'] = revised_staffing
            
            # Check if revised staffing violates threshold
            if revised_staffing < self.npt_threshold:
                logger.warning(f"Revised Staffing ({revised_staffing}) below threshold ({self.npt_threshold}) "
                             f"for {date} {workgroup} at {interval_time}")
        
        logger.info("Schedule_heatmap update completed")
    
    def update_manager_roster_meetings(self):
        """Update Manager_Roster with scheduled meetings"""
        logger.info("Updating Manager_Roster with scheduled meetings...")
        
        for manager_id, meetings in self.manager_meetings.items():
            if manager_id in self.manager_roster.index:
                # Group meetings by type
                meeting_times_by_type = defaultdict(list)
                
                for meeting in meetings:
                    meeting_type = meeting['meeting_type']
                    meeting_time = meeting['meeting_time'].strftime('%H:%M')
                    meeting_times_by_type[meeting_type].append(meeting_time)
                
                # Update manager roster columns
                for i, meeting_type in enumerate(self.input_data['Meeting Type'].head(4)):
                    if meeting_type in meeting_times_by_type:
                        col_name = ['P', 'Q', 'R', 'S'][i]
                        times = meeting_times_by_type[meeting_type]
                        self.manager_roster.loc[manager_id, col_name] = ', '.join(times)
        
        logger.info("Manager_Roster update completed")
    
    def generate_summary_report(self):
        """Generate comprehensive summary report"""
        logger.info("Generating summary report...")
        
        summary = {
            'total_meetings_scheduled': sum(len(meetings) for meetings in self.scheduled_meetings.values()),
            'meetings_by_type': defaultdict(int),
            'unscheduled_meetings': len(self.unscheduled_meetings),
            'team_huddle_distribution': dict(self.team_huddle_stats),
            'associates_scheduled': len([k for k, v in self.scheduled_meetings.items() if v]),
            'managers_involved': len([k for k, v in self.manager_meetings.items() if v])
        }
        
        # Count meetings by type
        for meetings in self.scheduled_meetings.values():
            for meeting in meetings:
                summary['meetings_by_type'][meeting['meeting_type']] += 1
        
        # Calculate Team Huddle distribution percentages
        huddle_percentages = {}
        for key, count in self.team_huddle_stats.items():
            if 'total' in key:
                base_key = key.replace('_total', '')
                first_count = self.team_huddle_stats.get(f"{base_key}_first", 0)
                second_count = self.team_huddle_stats.get(f"{base_key}_second", 0)
                
                if count > 0:
                    first_pct = (first_count / count) * 100
                    second_pct = (second_count / count) * 100
                    huddle_percentages[base_key] = {
                        'first_interval_pct': first_pct,
                        'second_interval_pct': second_pct,
                        'total_associates': count
                    }
        
        summary['huddle_percentages'] = huddle_percentages
        
        return summary
    
    def save_results(self):
        """Save results to output Excel file"""
        logger.info("Saving results to output file...")
        
        try:
            with pd.ExcelWriter(self.output_file_path, engine='openpyxl') as writer:
                # Save updated sheets
                self.input_data.to_excel(writer, sheet_name='Constraint', index=False)
                self.associate_roster.to_excel(writer, sheet_name='Associate_Roster', index=False)
                self.manager_roster.to_excel(writer, sheet_name='Manager_Roster', index=False)
                self.schedule_heatmap.to_excel(writer, sheet_name='Schedule_heatmap', index=False)
                
                # Add summary sheet
                summary = self.generate_summary_report()
                summary_df = pd.DataFrame([summary])
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
            
            logger.info(f"Results saved to {self.output_file_path}")
            
        except Exception as e:
            logger.error(f"Error saving results: {str(e)}")
            raise
    
    def run(self):
        """Execute the complete meeting scheduling process"""
        logger.info("Starting meeting scheduling process...")
        
        try:
            # Step 1: Load and prepare data
            self.load_data()
            
            # Step 2: Schedule Team Huddles
            self.schedule_team_huddles()
            
            # Step 3: Schedule One-to-One meetings
            self.schedule_one_to_one_meetings()
            
            # Step 4: Update Manager Roster
            self.update_manager_roster_meetings()
            
            # Step 5: Update Schedule Heatmap
            self.update_schedule_heatmap()
            
            # Step 6: Save results
            self.save_results()
            
            # Step 7: Print summary
            summary = self.generate_summary_report()
            self.print_summary(summary)
            
            logger.info("Meeting scheduling process completed successfully!")
            
        except Exception as e:
            logger.error(f"Error in scheduling process: {str(e)}")
            raise
    
    def print_summary(self, summary):
        """Print summary report to console"""
        print("\n" + "="*60)
        print("MEETING SCHEDULING SUMMARY REPORT")
        print("="*60)
        
        print(f"\n📊 OVERALL STATISTICS:")
        print(f"   Total meetings scheduled: {summary['total_meetings_scheduled']}")
        print(f"   Associates with meetings: {summary['associates_scheduled']}")
        print(f"   Managers involved: {summary['managers_involved']}")
        print(f"   Unscheduled meetings: {summary['unscheduled_meetings']}")
        
        print(f"\n📅 MEETINGS BY TYPE:")
        for meeting_type, count in summary['meetings_by_type'].items():
            print(f"   {meeting_type}: {count}")
        
        print(f"\n🤝 TEAM HUDDLE DISTRIBUTION:")
        for group, stats in summary['huddle_percentages'].items():
            print(f"   {group}:")
            print(f"      First interval: {stats['first_interval_pct']:.1f}%")
            print(f"      Second interval: {stats['second_interval_pct']:.1f}%")
            print(f"      Total associates: {stats['total_associates']}")
        
        if self.unscheduled_meetings:
            print(f"\n❌ UNSCHEDULED MEETINGS:")
            for meeting in self.unscheduled_meetings[:10]:  # Show first 10
                print(f"   Associate {meeting['associate_id']}: {meeting['meeting_type']} - {meeting['reason']}")
            if len(self.unscheduled_meetings) > 10:
                print(f"   ... and {len(self.unscheduled_meetings) - 10} more")
        
        print("\n" + "="*60)


def main():
    """Main function to run the meeting scheduler"""
    # File paths
    input_file = r"d:\babai\Consolidated.xlsx"
    output_file = r"d:\babai\Consolidated_Scheduled.xlsx"
    
    try:
        # Initialize and run scheduler
        scheduler = MeetingScheduler(input_file, output_file)
        scheduler.run()
        
        print(f"\n✅ Meeting scheduling completed successfully!")
        print(f"📄 Output file saved: {output_file}")
        
    except Exception as e:
        print(f"\n❌ Error in meeting scheduling: {str(e)}")
        logger.error(f"Main execution error: {str(e)}")


if __name__ == "__main__":
    main()