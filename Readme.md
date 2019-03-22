# Outlook Rule Manager

Imports and exports Outlook email rules to a simple JSON format.

## Caveats

Not all email rules are implemented. Check to make sure all components of your rules are represented in the exported JSON file before deleting them from Outlook.

## Usage

Export your current rules with the following command line:

```bat
RuleMgr.bat export RULEFILE.json
```

Import an existing `RULEFILE.json` with the following command line:

```bat
RuleMgr import RULEFILE.json
```

DELETE all of your current rules with the following command line:

```bat
RuleMgr.bat clear
```

## Advanced Usage

One way to maintain your own set of Outlook email rules is to create a new Git repository that references this project as a Git submodule. Assuming you have Git for Windows installed, paste the following batch script into a new `.bat` file and run it:

```bat
mkdir MyRules
cd MyRules
@git init
@git submodule add https://github.com/meustrus/OutlookRuleMgr.git OutlookRuleMgr
@git submodule init
@git submodule update
pushd OutlookRuleMgr
@git checkout release
popd
@echo @git submodule init > RuleMgr.bat
@echo @git submodule update >> RuleMgr.bat
@echo @call OutlookRuleMgr\RuleMgr.bat %%* >> RuleMgr.bat
@git add .
```

Then just `git commit`, `git remote add origin GIT_URL`, and `git push` to make your copy shareable.
