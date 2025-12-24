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
import importlib.metadata as importlib_metadata
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

def sanitize_filename(name):
    """Sanitize activity name for use in filename"""
    # Remove or replace invalid characters
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        name = name.replace(char, '-')
    # Replace spaces with underscores
    name = name.replace(' ', '_')
    # Limit length
    return name[:50] if name else 'Unnamed'

def download_gpx(client, activity_id, output_dir, date_str, activity_name=''):
    """Download activity as GPX format"""
    name_part = f'_{activity_name}' if activity_name else ''
    filename = output_dir / f'{date_str}{name_part}_{activity_id}.gpx'
    data = client.download_activity(activity_id, dl_fmt=client.ActivityDownloadFormat.GPX)
    filename.write_bytes(data)
    return filename

def download_tcx(client, activity_id, output_dir, date_str, activity_name=''):
    """Download activity as TCX format"""
    name_part = f'_{activity_name}' if activity_name else ''
    filename = output_dir / f'{date_str}{name_part}_{activity_id}.tcx'
    data = client.download_activity(activity_id, dl_fmt=client.ActivityDownloadFormat.TCX)
    filename.write_text(data.decode('utf-8'))
    return filename

def download_fit(client, activity_id, output_dir, date_str, activity_name=''):
    """Download and extract activity as FIT format"""
    name_part = f'_{activity_name}' if activity_name else ''
    zip_filename = output_dir / f'{date_str}{name_part}_{activity_id}.zip'
    data = client.download_activity(activity_id, dl_fmt=client.ActivityDownloadFormat.ORIGINAL)
    zip_filename.write_bytes(data)
    
    # Extract FIT file from ZIP and rename with date prefix
    final_filename = output_dir / f'{date_str}{name_part}_{activity_id}.fit'
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

def download_json(client, activity_id, output_dir, date_str, activity_name=''):
    """Download activity details as JSON format"""
    name_part = f'_{activity_name}' if activity_name else ''
    filename = output_dir / f'{date_str}{name_part}_{activity_id}.json'
    details = client.get_activity_details(activity_id)
    filename.write_text(json.dumps(details, indent=2))
    
    # Also save full activity data
    json_full = output_dir / f'{date_str}{name_part}_{activity_id}_full.json'
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
    safe_name = sanitize_filename(activity_name)
    
    print(f"{activity_name} ({activity_type}) - {start_time}")
    
    try:
        # Download based on format
        if format_type == 'gpx':
            filename = download_gpx(client, activity_id, output_dir, date_str, safe_name)
        elif format_type == 'tcx':
            filename = download_tcx(client, activity_id, output_dir, date_str, safe_name)
        elif format_type == 'fit':
            filename = download_fit(client, activity_id, output_dir, date_str, safe_name)
        elif format_type == 'json':
            filename = download_json(client, activity_id, output_dir, date_str, safe_name)
        
        print(f"  ✓ Saved to {filename.name}")
        return True
        
    except Exception as e:
        print(f"  ✗ Error: {e}")
        return False

def get_last_activity_date(output_dir, format_ext):
    """Get the date of the most recent activity file in the output directory"""
    try:
        files = list(output_dir.glob(f'*{format_ext}'))
        if not files:
            return None
        
        # Extract dates from filenames (format: YYYY-MM-DD_activityid.ext)
        dates = []
        for file in files:
            try:
                date_str = file.name.split('_')[0]
                date_obj = datetime.strptime(date_str, '%Y-%m-%d')
                dates.append(date_obj)
            except (ValueError, IndexError):
                continue
        
        return max(dates) if dates else None
    except Exception:
        return None

def main():
    parser = argparse.ArgumentParser(description='Export activities from Garmin Connect')
    parser.add_argument('--username', help='Garmin Connect username')
    parser.add_argument('--password', help='Garmin Connect password')
    parser.add_argument(
        '--tokenstore',
        help='Path to a tokenstore directory used by garminconnect to cache sessions (default: GARMIN_TOKENSTORE or ~/.garminconnect)',
    )
    parser.add_argument('-c', '--count', type=int, default=10, 
                        help='Number of recent activities to download (default: 10)')
    parser.add_argument('-f', '--format', choices=['gpx', 'tcx', 'fit', 'json'],  
                        default='fit', help='Export format (default: fit)')
    parser.add_argument('-d', '--directory', 
                        help='Output directory (default: GARMIN_OUTPUT_DIR from .env or ./garmin_exports)')
    parser.add_argument('--since-last', action='store_true', default=True,
                        help='Only download activities newer than the most recent file (default: enabled)')
    parser.add_argument('--all', action='store_true',
                        help='Download all activities, ignoring existing files')
    
    args = parser.parse_args()
    load_env()
    
    # If --all is specified, disable --since-last
    if args.all:
        args.since_last = False
    
    # Tokenstore (session cache directory)
    default_tokenstore = Path('~/.garminconnect').expanduser()
    tokenstore_path = Path(
        args.tokenstore
        or os.getenv('GARMIN_TOKENSTORE', str(default_tokenstore))
    ).expanduser()
    try:
        tokenstore_path.mkdir(parents=True, exist_ok=True)
    except Exception:
        # If we can't create the directory, we'll still try login without tokenstore later.
        pass

    # Get credentials
    username = args.username or os.getenv('GARMIN_USERNAME')
    password = args.password or os.getenv('GARMIN_PASSWORD')

    if not username and not tokenstore_path.exists():
        print("Error: Username required (via --username or GARMIN_USERNAME in .env)")
        print(f"Tokenstore directory not found at: {tokenstore_path}")
        return 1

    if username and not password and not tokenstore_path.exists():
        print("Error: Password required (via --password or GARMIN_PASSWORD in .env)")
        print(f"Tokenstore directory not found at: {tokenstore_path}")
        return 1
    
    # Set up output directory
    if args.directory:
        output_dir = Path(args.directory)
    else:
        output_dir = Path(os.getenv('GARMIN_OUTPUT_DIR', './garmin_exports'))
    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"Output directory: {output_dir}")
    
    # Check for last activity date if --since-last is used
    last_date = None
    if args.since_last:
        format_ext = f'.{args.format}'
        last_date = get_last_activity_date(output_dir, format_ext)
        if last_date:
            print(f"Last activity date found: {last_date.strftime('%Y-%m-%d')}")
            print(f"Will only download activities after this date")
        else:
            print("No existing activities found, downloading all requested activities")
    
    try:
        # Initialize Garmin client
        print("Connecting to Garmin Connect...")
        def prompt_mfa() -> str:
            return input('Enter Garmin MFA code: ').strip()

        client = Garmin(username, password, prompt_mfa=prompt_mfa)
        
        # Check if we have existing tokens to use tokenstore
        oauth1_token = tokenstore_path / 'oauth1_token.json'
        oauth2_token = tokenstore_path / 'oauth2_token.json'
        
        if oauth1_token.exists() or oauth2_token.exists():
            # Use cached tokens
            client.login(tokenstore=str(tokenstore_path))
        else:
            # First login - don't pass tokenstore, let library create tokens fresh
            client.login()
            # After successful login, save tokens for next time
            try:
                # Re-login with tokenstore to save the session
                client.login(tokenstore=str(tokenstore_path))
            except Exception:
                # If saving tokens fails, we're still authenticated, so continue
                pass
        
        print("✓ Successfully authenticated!")
        
        # Get activities
        print(f"\nFetching {args.count} most recent activities...")
        activities = client.get_activities(0, args.count)
        
        if not activities:
            print("No activities found")
            return 0
        
        print(f"Found {len(activities)} activities\n")
        
        # Filter activities if --since-last is used
        if last_date:
            original_count = len(activities)
            activities = [a for a in activities if datetime.strptime(
                a.get('startTimeLocal', '1970-01-01 00:00:00'), 
                '%Y-%m-%d %H:%M:%S'
            ) > last_date]
            filtered_count = original_count - len(activities)
            if filtered_count > 0:
                print(f"Filtered out {filtered_count} activities (already downloaded)\n")
            if not activities:
                print("No new activities to download")
                return 0
        
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
        try:
            gc_version = importlib_metadata.version('garminconnect')
        except Exception:
            gc_version = 'unknown'

        print("Common fixes:")
        print("- Confirm your Garmin username/password are still valid")
        print("- If Garmin enabled MFA on your account, re-run and enter the code when prompted")
        print(f"- Delete the tokenstore to force a fresh login: {tokenstore_path}")
        print(f"- Upgrade garminconnect (you currently have {gc_version}): pip3 install --upgrade garminconnect")
        return 1
        
    except GarminConnectConnectionError as e:
        print(f"\n✗ Connection error: {e}")
        return 1
        
    except Exception as e:
        print(f"\n✗ Unexpected error: {e}")
        return 1

if __name__ == '__main__':
    exit(main())
