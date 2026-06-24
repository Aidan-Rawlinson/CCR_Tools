# Project Instructions: [Project Name]

## Project Folder Path
This project maps to the following folder in the MCP root:

C:\mcp_projects\CCR_Tools

---

## Session Start Protocol

Trigger words: **Review** or **Wake up**

On receiving "Review" or "Wake up", read all three sources:
1. Project Files (held in Claude Desktop Project Files)
2. dynamic_docs folder (all four documents) -- check Current_State first to establish whether this is a new project or a project in progress before reading further
3. code_base folder (full structure with description of each component, but not the code itself)

If Current_State shows Status: New Project and code_base is empty, use exactly this format:

> Reviewed project documents.
> Reviewed dynamic_docs: Status is New Project.
> Reviewed code_base: empty -- full sight of project folder confirmed.
>[SHOW EMPTY LINE]
> Working on a secret project, are we, sir?
>[SHOW EMPTY LINE]
> Would you like me to talk through setting up Git on the folder?

If the user says yes, use the Git Setup Guide below as a strong reference. The guide assumes a typical Windows setup -- adapt your support to the user's actual situation rather than following the steps as a rigid script.

For all subsequent sessions, use this format:

> Reviewed project documents.
> Reviewed code_base: [brief summary of contents].
> Reviewed dynamic_docs: [what Current_State shows, what Next_Session flags].

---

## Session End Protocol

Trigger word: **Close-down**

Close-down should only be initiated once the session's work is complete and ready to commit. The sequence is:

1. Review all four dynamic_docs files (Current_State, Progression_Log, Decisions, Next_Session) and update any that need to reflect the session's work. Current_State in particular must not be left as "New Project" after a session has taken place -- if nothing else changed, the status alone must be updated.
2. Update `static_docs_mirror/project_files` to reflect the current state of the project -- write any new or changed files, remove any files that are no longer present in Project Files
3. If project documents are unchanged, state "No change to project documents" and move on -- no confirmation needed
4. If project documents have changed, prompt the user to update the Project Files in Claude Desktop and wait for confirmation
5. Verify the match between Project Files in context and `static_docs_mirror/project_files` -- list all files in the folder, compare against all files visible in Project Files, and flag any mismatch (missing file, extra file, or content difference) before proceeding
6. Surface everything in the project folder, flag anything that looks like it should not be committed, and ask the user to confirm or amend the .gitignore before proceeding
7. Present a suggested commit message using the Git convention -- a short summary title on the first line, a blank line, then a fuller description of what changed -- and wait for user approval
8. If the user approves, trigger the git_commit tool with the agreed message. If the user declines, discuss and revise before proceeding.
9. Confirm close-down is complete

---

## Scrap Session Protocol

Trigger word: **Scrap Session**

On receiving "Scrap Session":
1. Confirm with the user: "Are you sure? This will revert all changes made this session and cannot be undone."
2. Wait for user confirmation
3. Trigger the git_revert tool to restore the project folder to its last committed state
4. Respond: "Project folder restored -- close session"

---

## Git Setup Guide

This guide is a reference for helping the user set up Git on a new project folder. It assumes a typical Windows setup -- use it to shape your support, not as a script to follow rigidly. The user's environment may differ; adapt accordingly.

### Prerequisites
- Git installed on the machine
- A GitHub account

### Steps

**1. Open a command prompt**

**2. Navigate to the project folder**
```
cd C:\mcp_projects\[ProjectName]
```

**3. Initialise Git**
```
git init
```
Expected output: `Initialized empty Git repository in C:/mcp_projects/[ProjectName]/.git/`

**4. Create a repository on GitHub**
Create a new repository on GitHub and copy its URL (format: `https://github.com/[username]/[repo].git`)

**5. Link the remote**
```
git remote add origin https://github.com/[username]/[repo].git
```

**6. Set the default branch**
```
git branch -M main
```

**7. Stage and make the initial commit**
```
git add .
git commit -m "Initial commit"
```

**8. Set the upstream and push**
```
git push -u origin main
```

### Notes
- Steps 7 and 8 must be done in this order -- git push -u origin main will fail if there are no commits yet.
- From this point, all future commits and pushes are handled by the Close-down protocol.
- Authentication: GitHub requires a Personal Access Token (PAT) as the password when pushing via HTTPS. If prompted, use the PAT not the GitHub account password. Generate one at GitHub -- Settings -- Developer Settings -- Personal Access Tokens -- Tokens (classic) -- repo scope.

---

## Project-Specific Conventions

<!-- Add any project-specific conventions here: coding style, communication preferences, domain context, etc. -->
