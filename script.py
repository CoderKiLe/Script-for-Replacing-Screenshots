import os
import time
import subprocess
import pyautogui
import pygetwindow as gw
import psutil
import signal

def kill_process_tree(pid):
    """Kill a process and all its child processes"""
    try:
        parent = psutil.Process(pid)
        children = parent.children(recursive=True)
        
        # Kill children first
        for child in children:
            try:
                child.terminate()
            except psutil.NoSuchProcess:
                pass
        
        # Wait a moment for children to terminate
        gone, still_alive = psutil.wait_procs(children, timeout=3)
        
        # Force kill any remaining children
        for child in still_alive:
            try:
                child.kill()
            except psutil.NoSuchProcess:
                pass
        
        # Kill parent
        try:
            parent.terminate()
            parent.wait(3)  # Wait for parent to terminate
        except (psutil.NoSuchProcess, psutil.TimeoutExpired):
            try:
                parent.kill()
            except psutil.NoSuchProcess:
                pass
                
    except psutil.NoSuchProcess:
        pass

def process_single_project(project_dir):
    """Process a single project directory - runs dotnet run and captures screenshot"""
    print(f"--- Starting 'dotnet run' in {project_dir} ---")

    # Verify path exists
    if not os.path.exists(project_dir):
        print(f"Error: Directory not found: {project_dir}")
        return False

    # Verify it's a .NET project (look for .csproj files)
    csproj_files = [f for f in os.listdir(project_dir) if f.endswith('.csproj')]
    if not csproj_files:
        print(f"Warning: No .csproj files found in {project_dir}, skipping...")
        return False

    # --- STEP 0: Snapshot existing windows before launching ---
    print("--- Scanning existing windows... ---")
    existing_titles = set(gw.getAllTitles())
    process = None

    try:
        # Step 1: Run the dotnet application non-blocking
        print(f"Running: dotnet run in {project_dir}")
        process = subprocess.Popen(["dotnet", "run"], cwd=project_dir, shell=True)
        
        WAIT_TIME = 5
        print(f"--- Waiting {WAIT_TIME} seconds for application to load ---")
        time.sleep(WAIT_TIME)

        # --- STEP 2: Detect NEW windows ---
        current_titles = set(gw.getAllTitles())
        # Find titles that are in 'current' but were not in 'existing'
        new_titles = [t for t in (current_titles - existing_titles) if t.strip()]

        target_window = None

        if not new_titles:
            print("Error: No new windows detected after launch.")
            print("The application might have failed to start, or it didn't create a visible window.")
            if process:
                kill_process_tree(process.pid)
            return False
        
        elif len(new_titles) == 1:
            title = new_titles[0]
            print(f"--- Auto-detected application window: '{title}' ---")
            target_window = gw.getWindowsWithTitle(title)[0]
            
        else:
            print(f"--- Multiple new windows detected: {new_titles} ---")
            # Filter out system windows and choose the most likely candidate
            filtered_titles = [t for t in new_titles if not any(sys_word in t for sys_word in 
                            ['OleMainThreadWndName', 'MSCTFIME UI', 'Default IME'])]
            if filtered_titles:
                title = filtered_titles[0]
                print(f"--- Selected target window: '{title}' ---")
                target_window = gw.getWindowsWithTitle(title)[0]
            else:
                title = new_titles[0]
                print(f"--- Guessing target is: '{title}' ---")
                target_window = gw.getWindowsWithTitle(title)[0]

        # Step 3: Bring window to front
        if not target_window.isActive:
            try:
                target_window.activate()
                time.sleep(0.5) 
            except Exception as e:
                print(f"Warning: Could not force focus to window. Attempting capture anyway. ({e})")

        # Step 4: Take screenshot
        print("--- Capturing Screenshot ---")
        screenshot = pyautogui.screenshot(region=(
            target_window.left, 
            target_window.top, 
            target_window.width, 
            target_window.height
        ))

        # Step 5: Save with project name for uniqueness
        project_name = os.path.basename(project_dir)
        save_path = os.path.join(project_dir, f"{project_name}_screenshot.png")
        screenshot.save(save_path)

        print(f"Success! Screenshot saved to: {save_path}")

        # Step 6: Close the application window first
        print("--- Closing application window ---")
        try:
            # Try to close the window gracefully first
            target_window.close()
            time.sleep(1)
        except Exception as e:
            print(f"Warning: Could not close window gracefully: {e}")

        # Step 7: Kill the process tree
        print("--- Terminating dotnet process ---")
        if process:
            kill_process_tree(process.pid)
        
        # Additional cleanup - look for any remaining dotnet processes related to this project
        time.sleep(1)
        cleanup_stray_processes(project_dir)
        
        return True

    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        if process:
            kill_process_tree(process.pid)
        return False

def cleanup_stray_processes(project_dir):
    """Clean up any stray dotnet processes that might still be running"""
    try:
        project_name = os.path.basename(project_dir)
        for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
            try:
                # Look for dotnet processes that might be related to our project
                if proc.info['name'] and 'dotnet' in proc.info['name'].lower():
                    cmdline = proc.info['cmdline'] or []
                    if any(project_name in str(arg) for arg in cmdline):
                        print(f"Found stray process: {proc.info['pid']} - terminating")
                        proc.terminate()
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
    except Exception as e:
        print(f"Error during stray process cleanup: {e}")

def find_cs_projects(main_directory):
    """Find all CS project directories within the main directory structure"""
    cs_projects = []
    
    print(f"Scanning main directory: {main_directory}")
    
    # Walk through all directories under the main directory
    for root, dirs, files in os.walk(main_directory):
        # Check if this is a "CS" directory
        if os.path.basename(root).upper() == "CS":
            print(f"Found CS directory: {root}")
            # Look for project directories within the CS folder
            for item in os.listdir(root):
                item_path = os.path.join(root, item)
                if os.path.isdir(item_path):
                    # Check if it contains .csproj files
                    if any(f.endswith('.csproj') for f in os.listdir(item_path)):
                        cs_projects.append(item_path)
                        print(f"  Found C# project: {item_path}")
    
    return cs_projects

def search_resx_file(project_dir):
    """Search for a .resx file in the project directory"""
    print("--- Searching for .resx files ---")
    for root, dirs, files in os.walk(project_dir):
        for file in files:
            if file.endswith('.resx'):
                print(f"Found .resx file: {file} in {root}" )
                return os.path.join(root, file)
    return None


def run_for_all_projects():
    """Main function to process all projects"""
    # Get the main directory from user input
    MAIN_DIR = input("Enter the main directory path: ").strip().strip('"').strip("'")
    
    # Verify main directory exists
    if not os.path.exists(MAIN_DIR):
        print(f"Error: Main directory not found: {MAIN_DIR}")
        return
    
    # Find all CS projects
    projects = find_cs_projects(MAIN_DIR)
    
    if not projects:
        print("No C# projects found in the directory structure.")
        return
    
    print(f"\nFound {len(projects)} C# project(s) to process:")
    for i, project in enumerate(projects, 1):
        print(f"  {i}. {project}")
    
    print(f"\nStarting batch processing...")
    
    successful = 0
    failed = 0
    
    for i, project_dir in enumerate(projects, 1):
        print(f"\n{'='*60}")
        print(f"Processing project {i}/{len(projects)}: {os.path.basename(project_dir)}")
        print(f"{'='*60}")

        search_resx_file(project_dir=project_dir)

        if process_single_project(project_dir):
            successful += 1
        else:
            failed += 1
        
        # Small delay between projects
        time.sleep(3)
    
    print(f"\n{'='*60}")
    print(f"Batch processing completed!")
    print(f"Successful: {successful}")
    print(f"Failed: {failed}")
    print(f"Total: {len(projects)}")
    print(f"{'='*60}")

if __name__ == "__main__":
    # Ask user if they want to process single project or batch
    choice = input("Choose mode:\n1 - Single project (original behavior)\n2 - Batch process all projects in directory\nEnter choice (1 or 2): ").strip()
    
    if choice == "2":
        run_for_all_projects()
    else:
        # Original single project functionality
        PROJECT_DIR = input("Enter the full path to your project directory: ").strip().strip('"').strip("'")
        process_single_project(PROJECT_DIR)