# ProKnow Identity and Access Management Synchronization Tool

The ProKnow Identity and Access Management Synchronization Tool provides a Python script that can automatically synchronize Workspaces, Roles, and Users from a set of Excel workbooks to a ProKnow domain. This tool is focused on supporting large-scale implementations of ProKnow where managing the various permutations of Roles can be challenging.


## Getting Started

### Installing Python

Before you can run the synchronization script, you must first ensure that you have a recent version of [Python 3](https://www.python.org/downloads/) installed on your system. You can check the current version of Python by running the following command in a terminal:

```
$ python --version
```

If this command reports a version starting with "3" (e.g., "Python 3.9.5"), you have a suitable version of Python installed. Please note that in some older systems, the main `python` executable may be version "2". In that case, you may want to try to run the following command instead:

```
$ python3 --version
```

If command reports a version starting with "3" you should substitute any `python` calls in this document with `python3` instead.

### Installing Python Requirements

The synchronization scripts utilize several packages and modules that must be installed prior to executing the script. The easiest way to install these packages and modules in an isolated way is to utilize a [Python Virtual Environment](https://docs.python.org/3/tutorial/venv.html). Begin my running the following command from within the source directory:

```
$ python3 -m venv .venv
```

This will create a virtual environment in the `.venv` folder. Once you’ve created a virtual environment, you may activate it.

On Windows, run:
```
.venv\Scripts\activate.bat
```

On Unix or MacOS, run:
```
$ source .venv/bin/activate
```

Once activated, you may install the necessary packages and modules in the virtual environment by running the following command (note that the command is run from within the virtual environment):

```
(.venv) $ pip install -r requirements.txt
```

Once installed, you can then run the script as described below.

Please note that in most systems, activating a virtualenv gives you a shell function named:
```
$ deactivate
```
which you can run in order to exit the virtual environment and put things back to normal.


## Workspace, Role, and User Data

Before running the synchronization script, you must first provide the necessary Excel workbooks which contain the desired workspaces, roles, and user data. By default, this data must be stored within a single `data` directory, in the following structure:

```
data/
├─ users/
│  ├─ users1.xlsx
│  ├─ users2.xlsx
│  └─ usersN.xlsx
├─ roles.xlsx
└─ workspaces.xlsx
```

Please note that the names of expected files and directories (and the location of the data directory) may be customized via command line arguments.

Example Excel workbooks have been provided in the `examples` directory. However, please note that the synchronization script can be customized (via command line arguments) in order to utilize different columns than the ones used in the examples.


## Synchronization

Once you have installed the prerequisites and provided the necessary Excel workbooks containing the Workspace, Role, and User data, you may run the synchronization script by running the following command (substituting proper values for the url, credentials, and location of the data directory):

```
(.venv) $ python sync.py --url https://example.proknow.com --credentials /path/to/credentials.json ./data
```

Please note that you will be prompted each time before creating or updating records, and the script never deletes any records.

If you wish to customize the synchronization (for example, change the expected column names), you may run access the help information for the script by running the following command:

```
(.venv) $ python sync.py --help
```
