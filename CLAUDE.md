# Project Memory

## User Info
• Address user as Ben.
• Ben uses an ASUS Vivobook S 15 (Q5507Q) running Windows 11 Home with a Snapdragon X Plus ARM64 processor, 16 GB RAM, and IPv4 address 192.168.1.173.
• Ben uses WSL on Windows and his projects are stored in /home/benau/
• OpenAI API key is stored in the environmental variable OPENAI_API_KEY
• Ben prefers system-wide Python package installations in WSL rather than using virtual environments. Because WSL follows PEP 668, pip is blocked from installing packages system-wide by default. To work around this, Ben consistently uses the --break-system-packages flag with pip, for example: pip install package-name --break-system-packages. When a package is available through apt and the version is acceptable, he prefers to use sudo apt install for better integration with the system's package manager and dependency resolution. If a package is not available via apt or a newer version is needed, he uses sudo pip install --break-system-packages to install the latest version from PyPI. This approach aligns with his system preferences while minimizing conflicts and permission issues.

## Git Commands
When Ben says "git setup" or "setup git", run these commands:
1. `echo -e ".env\n__pycache__/\n*.pyc\ngitignore/" > .gitignore`
2. `git init -b main`
3. `git add .`
4. `git commit -m "Initial commit"`
5. `source ~/.bashrc`
6. `gh repo create $(basename $(pwd)) --public --source=. --remote=origin --push`

When Ben says "provide me update git command," respond with: `git add . && git commit -m "Describe your changes here" && git push`, including the appropriate project name or commit message as needed.

Ben stores large files and databases in a directory named `gitignore`, which is excluded via `.gitignore`.