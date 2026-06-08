import os
import zipfile

def package_project():
    release_filename = "vcb-exchange-rate-release.zip"
    exclude_dirs = {".venv", ".git", "__pycache__", "Backup", "scratch"}
    exclude_files = {
        release_filename,
        "vcb_rates.log",
        "temp_api_response.json",
        "temp_current_data.csv",
    }
    
    # Files/folders to explicitly include
    include_paths = ["scripts", "docs", "README.md", "requirements.txt", "LICENSE"]

    print(f"📦 Packaging project files into {release_filename}...")
    
    with zipfile.ZipFile(release_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk("."):
            # Skip excluded directories
            dirs[:] = [d for d in dirs if d not in exclude_dirs]
            
            for file in files:
                if file in exclude_files:
                    continue
                
                file_path = os.path.join(root, file)
                rel_path = os.path.relpath(file_path, ".")
                
                # Check if file belongs to included paths
                is_included = any(
                    rel_path.startswith(inc) or rel_path == inc 
                    for inc in include_paths
                )
                
                if is_included:
                    print(f"  + Adding: {rel_path}")
                    zipf.write(file_path, rel_path)
                    
    print(f"🎉 Packaging complete! Release archive saved as: {os.path.abspath(release_filename)}")

if __name__ == "__main__":
    package_project()
