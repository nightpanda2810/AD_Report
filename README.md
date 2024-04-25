# Starting application structure.

## Folder structure
* Project Root
    * app
        * main application files
    * extras
        * misc additional files
    * modules
        * other required files
    * testing
        * folder to hold test files, contents excluded from git

## Usage
1. Run to install pipreqs if not already installed.
    - pip install pipreqs
2. Run to install base requirements.
    - pip install -r requirements.txt
3. Run to keep Git history and begin a new repository.
    - git remote rm origin
4. Run to update the requirements file.
    - pipreqs --force