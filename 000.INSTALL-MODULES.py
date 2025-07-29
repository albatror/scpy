import subprocess

# List of packages to install
packages = ["tkinter", "os", "openpyxl", "pandas", "re"]

# Function to install a package
def install_package(package):
    subprocess.check_call(["pip", "install", package])

# Install each package
for package in packages:
    install_package(package)