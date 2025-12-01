#!/usr/bin/env python3
"""
Modern Garmin Connect Export Script
Uses the garminconnect library for authentication
"""

import os
import json
import argparse
import zipfile
from datetime import datetime
from pathlib import Path
from garminconnect import Garmin, GarminConnectAuthenticationError, GarminConnectConnectionError

def load_env():
    """Load environment variables from .env file if it exists"""
    env_file = Path(__file__).parent / '.env'
    if env_file.exists():
        with open(env_file, 'r') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    key, value = line.split('=', 1)
                    os.environ[key.strip()] = value.strip()

def get_date_string(start_time):
    """Parse activity start time and return formatted date string"""
    try:
        date_obj = datetime.strptime(start_time, '%Y-%m-%d %H:%M:%S')
        return date_obj.strftime('%Y-%m-%d')
    except:
        return 'unknown_date'

def download_gpx(client, activity_id, output_dir, date_str):
    """Download activity as GPX format"""
    filename = output_dir / f'{date_str}_{activity_id}.gpx'
    data = client.download_activity(activity_id, dl_fmt=client.ActivityDownloadFormat.GPX)
    filename.write_bytes(data)
    return filename

def download_tcx(client, activity_id, output_dir, date_str):
    """Download activity as TCX format"""
    filename = output_dir / f'{date_str}_{activity_id}.tcx'
    data = client.download_activity(activity_id, dl_fmt=client.ActivityDownloadFormat.TCX)
    filename.write_text(data.decode('utf-8'))
    return filename

def download_fit(client, activity_id, output_dir, date_str):
    """Download and extract activity as FIT format"""
    zip_filename = output_dir / f'{date_str}_{activity_id}.zip'
    data = client.download_activity(activity_id, dl_fmt=client.ActivityDownloadFormat.ORIGINAL)
    zip_filename.write_bytes(data)
    
    # Extract FIT file from ZIP and rename with date prefix
    final_filename = output_dir / f'{date_str}_{activity_id}.fit'
    try:
        with zipfile.ZipFile(zip_filename, 'r') as zip_ref:
            extracted_files = zip_ref.namelist()
            if extracted_files:
                # Extract the first file (usually the .fit file)
                zip_ref.extract(extracted_files[0], output_dir)
                extracted_path = output_dir / extracted_files[0]
                # Rename to date-prefixed filename
                extracted_path.rename(final_filename)
        zip_filename.unlink()  # Remove ZIP after extraction
    except zipfile.BadZipFile:
        # Not a ZIP, just rename to .fit
        zip_filename.rename(final_filename)
    
    return final_filename

def download_json(client, activity_id, output_dir, date_str):
    """Download activity details as JSON format"""
    filename = output_dir / f'{date_str}_{activity_id}.json'
    details = client.get_activity_details(activity_id)
    filename.write_text(json.dumps(details, indent=2))
    
    # Also save full activity data
    json_full = output_dir / f'{date_str}_{activity_id}_full.json'
    full_activity = client.get_activity(activity_id)
    json_full.write_text(json.dumps(full_activity, indent=2))
    
    return filename

def download_activity(client, activity, output_dir, format_type):
    """Download a single activity in the specified format"""
    activity_id = activity['activityId']
    activity_name = activity.get('activityName', 'Unnamed Activity')
    activity_type = activity.get('activityType', {}).get('typeKey', 'unknown')
    start_time = activity.get('startTimeLocal', 'unknown time')
    date_str = get_date_string(start_time)
    
    print(f"{activity_name} ({activity_type}) - {start_time}")
    
    try:
        # Download based on format
        if format_type == 'gpx':
            filename = download_gpx(client, activity_id, output_dir, date_str)
        elif format_type == 'tcx':
            filename = download_tcx(client, activity_id, output_dir, date_str)
        elif format_type == 'fit':
            filename = download_fit(client, activity_id, output_dir, date_str)
        elif format_type == 'json':
            filename = download_json(client, activity_id, output_dir, date_str)
        
        print(f"  ✓ Saved to {filename.name}")
        return True
        
    except Exception as e:
        print(f"  ✗ Error: {e}")
        return False

def main():
    parser = argparse.ArgumentParser(description='Export activities from Garmin Connect')
    parser.add_argument('--username', help='Garmin Connect username')
    parser.add_argument('--password', help='Garmin Connect password')
    parser.add_argument('-c', '--count', type=int, default=10, 
                        help='Number of recent activities to download (default: 10)')
    parser.add_argument('-f', '--format', choices=['gpx', 'tcx', 'fit', 'json'],  
                        default='fit', help='Export format (default: fit)')
    parser.add_argument('-d', '--directory', 
                        help='Output directory (default: GARMIN_OUTPUT_DIR from .env or ./garmin_exports)')
    
    args = parser.parse_args()
    load_env()
    
    # Get credentials
    username = args.username or os.getenv('GARMIN_USERNAME')
    password = args.password or os.getenv('GARMIN_PASSWORD')
    
    if not username or not password:
        print("Error: Username and password required (via --username/--password or .env file)")
        return 1
    
    # Set up output directory
    if args.directory:
        output_dir = Path(args.directory)
    else:
        output_dir = Path(os.getenv('GARMIN_OUTPUT_DIR', './garmin_exports'))
    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"Output directory: {output_dir}")
    
    try:
        # Initialize Garmin client
        print("Connecting to Garmin Connect...")
        client = Garmin(username, password)
        client.login()
        print("✓ Successfully authenticated!")
        
        # Get activities
        print(f"\nFetching {args.count} most recent activities...")
        activities = client.get_activities(0, args.count)
        
        if not activities:
            print("No activities found")
            return 0
        
        print(f"Found {len(activities)} activities\n")
        
        # Download each activity
        success_count = 0
        for i, activity in enumerate(activities, 1):
            print(f"[{i}/{len(activities)}] ", end='')
            if download_activity(client, activity, output_dir, args.format):
                success_count += 1
        
        print(f"\n✓ Export complete! {success_count}/{len(activities)} activities downloaded")
        return 0
        
    except GarminConnectAuthenticationError as e:
        print(f"\n✗ Authentication failed: {e}")
        print("Please check your username and password")
        return 1
        
    except GarminConnectConnectionError as e:
        print(f"\n✗ Connection error: {e}")
        return 1
        
    except Exception as e:
        print(f"\n✗ Unexpected error: {e}")
        return 1

if __name__ == '__main__':
    exit(main())
