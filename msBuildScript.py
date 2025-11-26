import os
import time
import subprocess
import pyautogui
import pygetwindow as gw
import psutil
import logging
from pathlib import Path
import signal
from resx_ico_replace import ResxIconUpdater
import xml.etree.ElementTree as ET


# Create the logger
logger = logging.getLogger("MsBuildScript Log")
logger.setLevel(logging.DEBUG)

# Handler for DEBUG logs
errorLoggingHandler = logging.FileHandler("error.log", encoding="utf-8")
errorLoggingHandler.setLevel(logging.ERROR)
errorLoggingHandler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))

# Handler for INFO logs
infoLoggingHandler = logging.FileHandler("debug.log", encoding="utf-8")
infoLoggingHandler.setLevel(logging.DEBUG)
infoLoggingHandler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))

# Attach handlers to the logger
logger.addHandler(errorLoggingHandler)
logger.addHandler(infoLoggingHandler)


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
                logger.debug(f"Child process {child.pid} already terminated.")
                pass
        
        # Wait a moment for children to terminate
        gone, still_alive = psutil.wait_procs(children, timeout=3)
        
        # Force kill any remaining children
        for child in still_alive:
            try:
                child.kill()
            except psutil.NoSuchProcess as e:
                logger.debug(f"Child process {child.pid} already terminated: {e}")
                pass
        
        # Kill parent
        try:
            parent.terminate()
            parent.wait(3)  # Wait for parent to terminate
        except (psutil.NoSuchProcess, psutil.TimeoutExpired):
            try:
                parent.kill()
            except psutil.NoSuchProcess as e:
                logger.debug(f"Process {pid} already terminated: {e}")
                pass
                
    except psutil.NoSuchProcess as e:
        logger.debug(f"Process {pid} already terminated: {e}")
        pass

def run_subprocess(command, cwd, wait = True, debug_name = "Subprocess"):
    """Run a subprocess command with timeout and return process"""
    try:
        process = subprocess.Popen(
            command, 
            cwd=cwd, 
            shell=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )
        
        if(not wait):
            print(f"{debug_name}[{command}] started with PID: {process.pid}")
            return process
        
        stdout, stderr = process.communicate()
        logger.debug(f"[{cwd}]-{debug_name}[{command}] stdout: {stdout.decode()}")

        if wait:
            process.wait()

        if process.returncode != 0:
            logger.error(f"{debug_name}[{command}] failed with return code: {process.returncode}")
            print(f"[{cwd}]-{debug_name} failed with return code: {process.returncode}")
            
            if stderr:
                logger.error(f"[{cwd}]-{debug_name}[{command}] errors: {stderr.decode()}")
                print(f"{debug_name}[{command}] errors: {stderr.decode()}")
                raise Exception(stderr.decode())
            raise Exception(f"{debug_name}[{command}] failed with return code: {process.returncode}")
        print(f"{debug_name} successful!")
        logger.debug(f"{debug_name} successful!")
    except Exception as e:
        logger.error(f"[{cwd}]-{debug_name}[{command}] error: {e}")
        print(f"{debug_name}[{command}] error: {e}")
        raise



def build_and_run_netframework_project(project_dir, csproj):
    """Build and run .NET Framework project"""    
    print(f"Building project: {csproj}")

    # Step 1: Clean the project
    run_subprocess(f'dotnet clean {csproj}', cwd=project_dir, debug_name="Clean")

    # Step 2: Restore packages
    run_subprocess(f'dotnet restore {csproj}', cwd=project_dir, debug_name="Restore")

    # Step 3: Build the project
    print(f"Building project: {csproj}")
    run_subprocess(f'msbuild {csproj}', cwd=project_dir, debug_name="Build")

    # Step 4: Dotnet run
    return run_subprocess(f'dotnet run {csproj}', cwd=project_dir, wait=False, debug_name="Run")


def find_output_executable(project_dir, csproj_filename):
    """Find the output executable for a .NET Framework project"""
    project_name = os.path.splitext(csproj_filename)[0]
    
    # Common output paths for .NET Framework projects
    possible_paths = [
        os.path.join(project_dir, "bin", "Debug", project_name + ".exe"),
        os.path.join(project_dir, "bin", "Release", project_name + ".exe"),
        os.path.join(project_dir, "bin", "Debug", project_name + ".vshost.exe"),
        os.path.join(project_dir, "bin", "Release", project_name + ".vshost.exe"),
    ]
    
    for path in possible_paths:
        if os.path.exists(path):
            return path
    
    # If not found, try to find any .exe in bin directories
    bin_debug_path = os.path.join(project_dir, "bin", "Debug")
    bin_release_path = os.path.join(project_dir, "bin", "Release")
    
    for bin_path in [bin_debug_path, bin_release_path]:
        if os.path.exists(bin_path):
            for file in os.listdir(bin_path):
                if file.endswith('.exe') and not file.endswith('.vshost.exe'):
                    return os.path.join(bin_path, file)
    
    return None

def detect_new_window(existing_titles, retry_count = 0, max_retries = 20):
    if retry_count >= max_retries:
        print("Error: Maximum retries reached while detecting new window.")
        raise Exception("Maximum retries reached while detecting new window.")
    
    retry_count += 1
    current_titles = set(gw.getAllTitles())
    # Find titles that are in 'current' but were not in 'existing'
    new_titles = [t for t in (current_titles - existing_titles) if t.strip()]

    target_window = None

    if not new_titles:
        time.sleep(0.2)
        return detect_new_window(existing_titles, retry_count, max_retries)
    
    # Filter out irrelevant windows
    filtered_titles = [
        t for t in new_titles 
        if t and not any(sys_word in t for sys_word in 
            ['OleMainThreadWndName', 'MSCTFIME UI', 'Default IME', 'ConsoleWindowClass'])
    ]
    
    if filtered_titles:
        title = filtered_titles[0]
        print(f"--- Selected application window: '{title}' ---")
        try:
            target_window = gw.getWindowsWithTitle(title)[0]
        except IndexError:
            print(f"Window with title '{title}' not found, trying first available...")
            if new_titles:
                target_window = gw.getWindowsWithTitle(new_titles[0])[0]
            else:
                print("No windows found")
    else:
        print("No suitable windows found after filtering")

    if target_window is not None:
        return target_window
    
    time.sleep(0.2)
    return detect_new_window(existing_titles, retry_count, max_retries)

def bring_window_to_front_take_screenshot(target_window, csproj):
    if target_window is None:
        logger.debug("No target window to bring to front.")
        return
    
    if not target_window.isActive:
        try:
            target_window.activate()
            time.sleep(1)  # Give more time for window activation
        except Exception as e:
            logger.error(f"Warning: Could not force focus to window. Attempting capture anyway. ({e})") 
            print(f"Warning: Could not force focus to window. Attempting capture anyway. ({e})")

    target_window.maximize()
    time.sleep(2)
    
    # Step 4: Take screenshot
    print("--- Capturing Screenshot ---")
    try:
        screenshot = pyautogui.screenshot(region=(
            target_window.left, 
            target_window.top, 
            target_window.width, 
            target_window.height
        ))
        # Step 5: Save with project name for uniqueness
        project_name = os.path.basename(csproj)
        save_path = os.path.join(csproj, f"{project_name}_screenshot.png")
        screenshot.save(save_path)
        logger.debug(f"Screenshot saved to: {save_path}")

    except Exception as e:
        logger.error(f"Failed to capture screenshot: {e}")    
        print(f"Failed to capture screenshot: {e}")
        return

def close_application(target_window):
    print("--- Closing application ---")
    if target_window:
        try:
            # Try to close the window gracefully first
            target_window.close()
            logger.debug("Sent close signal to application window.")    
            time.sleep(2)
        except Exception as e:
            logger.error(f"Warning: Could not close window gracefully: {e}")    
            print(f"Warning: Could not close window gracefully: {e}")
    return 

def process_single_project(project_dir):
    """Process a single project directory - runs the application and captures screenshot"""
    print(f"--- Processing project in {project_dir} ---")
    logger.debug(f"Processing project in {project_dir}")
    # Verify path exists
    if not os.path.exists(project_dir):
        logger.error(f"Directory not found: {project_dir}") 
        print(f"Error: Directory not found: {project_dir}")
        return False

    # Verify it's a .NET project (look for .csproj files)
    csproj_files = [f for f in os.listdir(project_dir) if f.endswith('.csproj')]
    if not csproj_files:
        print(f"Warning: No .csproj files found in {project_dir}, skipping...")
        logger.error(f"No .csproj files found in {project_dir}, skipping...") 
        return False

    # --- STEP 0: Snapshot existing windows before launching ---
    print("--- Scanning existing windows... ---")
    existing_titles = set(gw.getAllTitles())
    logger.debug(f"Existing window titles: {existing_titles}")
    process = None
    try:
        # Step 1: Try to run the .NET Framework project
        print(f"Attempting to build and run .NET Framework projects...")
        logger.debug(f"Attempting to build and run .NET Framework projects...")
        csproj_files = [f for f in os.listdir(project_dir) if f.endswith('.csproj')]
        
        if not csproj_files:
            logger.error(f"No .csproj files found in {project_dir}, skipping...")
            return None
        
        for csproj in csproj_files:
            app_process = None
            try:
                ResxIconUpdater('C1.ico').search_and_update(project_dir)
                app_process = build_and_run_netframework_project(project_dir, csproj)

                print("--- Detecting new application window... ---" )
                logger.debug(f"[{csproj}]-Detecting new application window... ---" )
                target_window = detect_new_window(existing_titles)
                
                if target_window is None:
                    logger.error(f"Could not detect application window for project {csproj}, skipping screenshot...")
                    print(f"Could not detect application window for project {csproj}, skipping screenshot...")
                    continue

                bring_window_to_front_take_screenshot(target_window, project_dir)
                print("--- Closing application... ---")
                close_application(target_window)
            except Exception as e:
                logger.error(f"[{csproj}][{project_dir}]-Build/Run failed for {csproj}: {e}")   
                print(f"Build/Run failed for {csproj}: {e}")
                continue
            finally:
                print("--- Cleaning up stray processes... ---")
                logger.debug(f"[{project_dir}]-Cleaning up stray processes...")
                if(app_process is not None):
                    logger.debug(f"[{project_dir}]-Killing process tree for PID: {app_process.pid}")    
                    print("--- Killing process tree...---")
                    kill_process_tree(app_process.pid)
        return True      
    except Exception as e:
        logger.error(f"[{project_dir}]-An unexpected error occurred: {e}")
        print(f"An unexpected error occurred: {e}")
        if process:
            logger.debug(f"[{project_dir}]-Killing process tree for PID: {process.pid}")
            kill_process_tree(process.pid)
        return False

def cleanup_stray_processes(project_dir):
    """Clean up any stray processes that might still be running"""
    try:
        project_name = os.path.basename(project_dir)
        for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
            try:
                # Look for processes that might be related to our project
                cmdline = proc.info['cmdline'] or []
                if any(project_name in str(arg) for arg in cmdline):
                    logger.debug(f"Found stray process related to project: {proc.info['pid']} - terminating")
                    print(f"Found stray process related to project: {proc.info['pid']} - terminating")
                    proc.terminate()
            except (psutil.NoSuchProcess, psutil.AccessDenied) as e:
                logger.error(f"Error accessing process info: {e}")

    except Exception as e:
        logger.error(f"Error during stray process cleanup: {e}")    
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

def run_for_all_projects():
    """Main function to process all projects"""
    # Get the main directory from user input
    MAIN_DIR = input("Enter the main directory path: ").strip().strip('"').strip("'")
    
    # Verify main directory exists
    if not os.path.exists(MAIN_DIR):
        logger.error(f"Main directory not found: {MAIN_DIR}")
        print(f"Error: Main directory not found: {MAIN_DIR}")
        return
    
    # Find all CS projects
    projects = find_cs_projects(MAIN_DIR)
    
    if not projects:
        logger.error("No C# projects found in the directory structure.")
        print("No C# projects found in the directory structure.")
        return

    print(f"\nFound {len(projects)} C# project(s) to process:")
    logger.debug(f"Found {len(projects)} C# project(s) to process:")
    
    for i, project in enumerate(projects, 1):
        print(f"  {i}. {project}")
    
    print(f"\nStarting batch processing...")
    logger.debug("Starting batch processing...")
    
    successful = 0
    failed = 0
    
    for i, project_dir in enumerate(projects, 1):
        print(f"\n{'='*60}")
        print(f"Processing project {i}/{len(projects)}: {os.path.basename(project_dir)}")
        print(f"{'='*60}")
        
        try:
            if process_single_project(project_dir):
                successful += 1
            else:
                failed += 1
        except Exception as e:
            failed += 1
            continue
        
        # Longer delay between projects to ensure clean shutdown
        time.sleep(5)
    
    print(f"\n{'='*60}")
    print(f"Batch processing completed!")
    print(f"Successful: {successful}")
    print(f"Failed: {failed}")
    print(f"Total: {len(projects)}")
    print(f"{'='*60}")
    logger.debug(f"Batch processing completed! Successful: {successful}, Failed: {failed}, Total: {len(projects)}")


def exit_gracefully(signum, frame):
    logger.debug("Received termination signal. Exiting gracefully...")
    print("\nReceived termination signal. Exiting gracefully...")
    exit(0)

if __name__ == "__main__":
    signal.signal(signal.SIGTERM, exit_gracefully)
    # Ask user if they want to process single project or batch
    choice = input("Choose mode:\n1 - Single project (original behavior)\n2 - Batch process all projects in directory\nEnter choice (1 or 2): ").strip()
    logger.debug(f"Script started in mode: {choice}")

    if choice == "2":
        run_for_all_projects()
    else:
        # Original single project functionality
        PROJECT_DIR = input("Enter the full path to your project directory: ").strip().strip('"').strip("'")

        PROJECT_DIR = Path(r"K:\Source Clone Items\Winforms Code base (Samples)-Source Clone\NetFramework\Barcode\CS\BarcodeDemo")
        # PROJECT_DIR = "K:\Source Clone Items\Winforms Code base (Samples)-Source Clone\NetFramework\Barcode\CS\BarcodeDemo".strip().strip('"').strip("'")
        logger.debug(f"Processing single project: {PROJECT_DIR}")
        logger.error("This is just for testing purposes")
        process_single_project(PROJECT_DIR)
