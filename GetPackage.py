import os

# Get current working directory
cwd = os.getcwd()

# Run nuget command to install package in current directory
os.system(f'nuget install Agito.AAComm -Source "https://api.nuget.org/v3/index.json" -OutputDirectory {cwd}')