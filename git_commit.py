import os
import git
from datetime import datetime

# Specify the project folder path
script_directory = os.path.dirname(os.path.abspath(__file__))

# Initialize a Git repository object
repo = git.Repo(script_directory)

# Add all changes to the staging area
repo.git.add('--all')
print('Added everything to stage.')

# Create a commit with the current date and time as the message
commit_message = f'Automatic commit, {datetime.now()}'
repo.git.commit('-m', commit_message)
print('Commited everything.')

# Push the changes to the repository
repo.remotes.origin.push()
print('Pushed everything.')
