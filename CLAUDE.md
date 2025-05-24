# Project Memory

## Git Commands
When Ben says "git setup" or "setup git", run these commands:
1. `echo -e ".env\n__pycache__/\n*.pyc\ngitignore/" > .gitignore`
2. `git init -b main`
3. `git add .`
4. `git commit -m "Initial commit"`
5. `source ~/.bashrc`
6. `gh repo create project-name --public --source=. --remote=origin --push`

Ben stores large files and databases in a directory named `gitignore`, which is excluded via `.gitignore`.