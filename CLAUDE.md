# CLAUDE.md
This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview
This is a project to create scripts and tools for managing Creo Parametric model parameters in a user friendly manner. Creo is capable of importing and exporting lists of parameters in XML format. A custom sorting configuration has been created and exported (rp_config.xml) that filters the parameters table of an assembly, its subassemblies and parts to create a table containing only the parameters that need to be modified. This table was exported as example.xml and added to the repo. For convenience we will be refering to parts (.prt) and assemblies (.asm) as "CAD objects"

The xml file contains parameter fields CAGE_CODE, DESCRIPTION_1, DESCRIPTION_2, PART_NUMBER, and PTC_WM_NAME. It is sorted using a multilevel sort, first by CAD object and then alphabetically by parameter name. There is no other indication of which CAD object the parameter belongs to besides the sorting method; it is critical that this order be retained when exporting back into xml format to prevent incorrect parameter assignment. PTC_WM_NAME contains the name of the object, and the three fields above it (CAGE_CODE, DESCRIPTION_1, PART_NUMBER) are the parameters belonging to it. I have prepopulated one of the objects in example.xml:

```xml
    <Parameter Name="CAGE_CODE">
        <DataType>String</DataType>
        <Value>0AEX9</Value>
        <Description>*Design Activity CAGE Code</Description>
    </Parameter>
    <Parameter Name="DESCRIPTION_1">
        <DataType>String</DataType>
        <Value>JIC x JIC x FNPT Female Branch Tee 3/4"</Value>
        <Description>*Name/Description Line 1</Description>
    </Parameter>
    <Parameter Name="DESCRIPTION_2">
        <DataType>String</DataType>
        <Value></Value>
        <Description>Name/Description Line 2</Description>
    </Parameter>
    <Parameter Name="PART_NUMBER">
        <DataType>String</DataType>
        <Value>J12TTF</Value>
        <Description>*Part Number</Description>
    </Parameter>
    <Parameter Name="PTC_WM_NAME">
        <DataType>String</DataType>
        <Value>ssp-j12ttf_brnch_tee_12.prt</Value>
        <Access>Locked</Access>
```

From this we gather that the part ssp-j12ttf_brnch_tee_12.prt has the following parameters:
CAGE_CODE = 0AEX9
DESCRIPTION_1 = JIC x JIC x FNPT Female Branch Tee 3/4 
DESCRIPTION_2 = <null>
PART_NUMBER = J12TTF
PTC_WM_NAME = ssp-j12ttf_brnch_tee_12.prt

## Deliverables
Initially we will be creating the following scripts:
### param_xml_to_xls
    Takes an xml file exported from creo and converts it to a spreadsheet that sorts the parameters into a table with the following columns (in this order)
    PTC_WM_NAME | CAGE_CODE | PART_NUMBER | DESCRIPTION_1 | DESCRIPTION_2
    
    PTC_WM_NAME should not be editable

    All fields should be explicitly formatted as text.

    Tool should open a dialog box that allows the selection of an xml file and create an xls file or sheet with the same name.

### param_xls_to_xml
    Takes the xls file and re-sorts it back into a properly formatted xml file that retains the object>parameter-name order required for proper import
    Checks the original xml file to ensure there are no errors (covered in edge cases)
    The new xml file should have the same name as the original concatenated with the string "_FILLED"
    
## Edge Cases
Some possible edge cases that need to be considered, and the method to handle them
### Blank fields
    Many fields may be blank when the xml is exported from creo, with only PTC_WM_NAME being assured to be filled as it relates to the file name. When converting to xls these should appear as blank cells.
    When exporting back to xml from xls, only DESCRIPTION_2 may be left blank. Attempting to export an xls file with illegal blank fields should spit an error that requires the user to confirm or cancel.
    If a cell is left blank when exporting to xml (either an allowed field like DESCRIPTION_2 or from the user aknowledging the error and continuing anyways), populate that parameter in the xml but use the <Value></Value> format
    
### New fields or lines not in the original xml file
    It may occur that the user accidentally changes the column names or re-orders them, the tool should avoid this edge case completely by matching the columns by number in order, not name. If a mismatch occurs just spit an warning
    The tool will also check that the number of values in the spreadsheet matches the number of parameters in the original xml file. This accounts for the user accidentally adding or removing rows and should spit a warning if a mismatch occurs
    The tool should lock the first row and column of the sheet to avoid this altogether

### Mismatched PTC_WM_NAME
    This should not occur due to the lock on the first column but check it and spit a warning anyways

## Execution Strategy
The first step will be to evaluate the feasibility of this plan and make adjustments as necessary. The information here takes priority over the above sections as this is specific implementation detail. If the information here conflicts with that in the deliverable section, use the information in the execution strategy section:

I would like to have a macro enabled spreadsheet with an import and export button on the first sheet. The first sheet would have a short readme, an import button, an export button, and two list objects:
1. Lists the xml files in the same directory as the spreadsheet, sorted by creation date newest-first if possible or just alphabetically otherwise.
2. Lists the sheets of the workbook excluding the Manager (first) sheet, sorted newest-first (should match datetime stamps)

The lists will both be single-selection list boxes, meaning only one file or sheet can be selected at a time. When the 'import' button is pressed, the sheet will grab the xml file selected in the list box, and convert it to a workbook sheet named using the filename and current datetime stamp with the format <xml_file_name>&&"-"&&<yyyymmdd_hhmmss>&&".xml" ( This isnt actual vba code just a representation of the plan. I am using && here to denote concatenation, <> for dynamically generated strings, and "" for static strings)

The newly created sheet will show up in the sheet list box. 

When the 'export' button is pressed, whichever sheet is selected is converted back into an xml file with error checking. The new xml file is named with the same name as the sheet it was created from. It will appear in the xml list box

### Potential pitfalls
1. We need to make sure that we can process the xml file in vba
2. Do we need form objects or activex objects?
3. Will we be able to list the files in the spreadsheets directory in a list object?
4. Will the lists update automatically or will we need to add a refresh button? The sheets list can be updated on import of an xml file, but the xml file will need to update when exporting a sheet to xml (easy) or a new file is placed in the directory (hard, can happen any time; may need manual buttom)

## AI Guidance Boilerplate Below (Not Project-Specific)

* Ignore GEMINI.md and GEMINI-*.md files
* To save main context space, for code searches, inspections, troubleshooting or analysis, use code-searcher subagent where appropriate - giving the subagent full context background for the task(s) you assign it.
* ALWAYS read and understand relevant files before proposing code edits. Do not speculate about code you have not inspected. If the user references a specific file/path, you MUST open and inspect it before explaining or proposing fixes. Be rigorous and persistent in searching code for key facts. Thoroughly review the style, conventions, and abstractions of the codebase before implementing new features or abstractions.
* After receiving tool results, carefully reflect on their quality and determine optimal next steps before proceeding. Use your thinking to plan and iterate based on this new information, and then take the best next action.
* After completing a task that involves tool use, provide a quick summary of what you've done.
* For maximum efficiency, whenever you need to perform multiple independent operations, invoke all relevant tools simultaneously rather than sequentially.
* Before you finish, please verify your solution
* Do what has been asked; nothing more, nothing less.
* NEVER create files unless they're absolutely necessary for achieving your goal.
* ALWAYS prefer editing an existing file to creating a new one.
* NEVER proactively create documentation files (*.md) or README files. Only create documentation files if explicitly requested by the User.
* If you create any temporary new files, scripts, or helper files for iteration, clean up these files by removing them at the end of the task.
* When you update or modify core context files, also update markdown documentation and memory bank
* When asked to commit changes, exclude CLAUDE.md and CLAUDE-*.md referenced memory bank system files from any commits. Never delete these files.

<investigate_before_answering>
Never speculate about code you have not opened. If the user references a specific file, you MUST read the file before answering. Make sure to investigate and read relevant files BEFORE answering questions about the codebase. Never make any claims about code before investigating unless you are certain of the correct answer - give grounded and hallucination-free answers.
</investigate_before_answering>

<do_not_act_before_instructions>
Do not jump into implementatation or changes files unless clearly instructed to make changes. When the user's intent is ambiguous, default to providing information, doing research, and providing recommendations rather than taking action. Only proceed with edits, modifications, or implementations when the user explicitly requests them.
</do_not_act_before_instructions>

<use_parallel_tool_calls>
If you intend to call multiple tools and there are no dependencies between the tool calls, make all of the independent tool calls in parallel. Prioritize calling tools simultaneously whenever the actions can be done in parallel rather than sequentially. For example, when reading 3 files, run 3 tool calls in parallel to read all 3 files into context at the same time. Maximize use of parallel tool calls where possible to increase speed and efficiency. However, if some tool calls depend on previous calls to inform dependent values like the parameters, do NOT call these tools in parallel and instead call them sequentially. Never use placeholders or guess missing parameters in tool calls.
</use_parallel_tool_calls>

## Memory Bank System

This project uses a structured memory bank system with specialized context files. Always check these files for relevant information before starting work:

### Core Context Files

* **CLAUDE-activeContext.md** - Current session state, goals, and progress (if exists)
* **CLAUDE-patterns.md** - Established code patterns and conventions (if exists)
* **CLAUDE-decisions.md** - Architecture decisions and rationale (if exists)
* **CLAUDE-troubleshooting.md** - Common issues and proven solutions (if exists)
* **CLAUDE-config-variables.md** - Configuration variables reference (if exists)
* **CLAUDE-temp.md** - Temporary scratch pad (only read when referenced)

**Important:** Always reference the active context file first to understand what's currently being worked on and maintain session continuity.

### Memory Bank System Backups

When asked to backup Memory Bank System files, you will copy the core context files above and @.claude settings directory to directory @/path/to/backup-directory. If files already exist in the backup directory, you will overwrite them.

## Claude Code Official Documentation

When working on Claude Code features (hooks, skills, subagents, MCP servers, etc.), use the `claude-docs-consultant` skill to selectively fetch official documentation from docs.claude.com.

## Project Overview



## ALWAYS START WITH THESE COMMANDS FOR COMMON TASKS

**Task: "List/summarize all files and directories"**

```bash
fd . -t f           # Lists ALL files recursively (FASTEST)
# OR
rg --files          # Lists files (respects .gitignore)
```

**Task: "Search for content in files"**

```bash
rg "search_term"    # Search everywhere (FASTEST)
```

**Task: "Find files by name"**

```bash
fd "filename"       # Find by name pattern (FASTEST)
```

### Directory/File Exploration

```bash
# FIRST CHOICE - List all files/dirs recursively:
fd . -t f           # All files (fastest)
fd . -t d           # All directories
rg --files          # All files (respects .gitignore)

# For current directory only:
ls -la              # OK for single directory view
```

### BANNED - Never Use These Slow Tools

* ❌ `tree` - NOT INSTALLED, use `fd` instead
* ❌ `find` - use `fd` or `rg --files`
* ❌ `grep` or `grep -r` - use `rg` instead
* ❌ `ls -R` - use `rg --files` or `fd`
* ❌ `cat file | grep` - use `rg pattern file`

### Use These Faster Tools Instead

```bash
# ripgrep (rg) - content search 
rg "search_term"                # Search in all files
rg -i "case_insensitive"        # Case-insensitive
rg "pattern" -t py              # Only Python files
rg "pattern" -g "*.md"          # Only Markdown
rg -1 "pattern"                 # Filenames with matches
rg -c "pattern"                 # Count matches per file
rg -n "pattern"                 # Show line numbers 
rg -A 3 -B 3 "error"            # Context lines
rg " (TODO| FIXME | HACK)"      # Multiple patterns

# ripgrep (rg) - file listing 
rg --files                      # List files (respects •gitignore)
rg --files | rg "pattern"       # Find files by name 
rg --files -t md                # Only Markdown files 

# fd - file finding 
fd -e js                        # All •js files (fast find) 
fd -x command {}                # Exec per-file 
fd -e md -x ls -la {}           # Example with ls 

# jq - JSON processing 
jq. data.json                   # Pretty-print 
jq -r .name file.json           # Extract field 
jq '.id = 0' x.json             # Modify field
```

### Search Strategy

1. Start broad, then narrow: `rg "partial" | rg "specific"`
2. Filter by type early: `rg -t python "def function_name"`
3. Batch patterns: `rg "(pattern1|pattern2|pattern3)"`
4. Limit scope: `rg "pattern" src/`

### INSTANT DECISION TREE

```
User asks to "list/show/summarize/explore files"?
  → USE: fd . -t f  (fastest, shows all files)
  → OR: rg --files  (respects .gitignore)

User asks to "search/grep/find text content"?
  → USE: rg "pattern"  (NOT grep!)

User asks to "find file/directory by name"?
  → USE: fd "name"  (NOT find!)

User asks for "directory structure/tree"?
  → USE: fd . -t d  (directories) + fd . -t f  (files)
  → NEVER: tree (not installed!)

Need just current directory?
  → USE: ls -la  (OK for single dir)
```
