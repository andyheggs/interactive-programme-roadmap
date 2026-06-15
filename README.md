# Interactive Programme Roadmap

React + TypeScript + Vite single-page app for importing Microsoft Project XML and turning it into an interactive programme roadmap and integrated master schedule.

## What It Does

- Imports Microsoft Project XML locally in the browser.
- Parses project metadata, tasks, milestones, hierarchy, baselines, dependencies, resources, assignments and custom fields.
- Resolves readable custom field values for Stream, Roadmap Milestone, Milestone Type, Approval Body, Version, Visibility and Roadmap View.
- Provides Programme Roadmap, Integrated Schedule, Milestones, Governance, Delivery and Version / Release views.
- Filters by stream, roadmap view, milestone type, approval body, version, visibility, status, critical items, roadmap milestones and delays.
- Shows baseline finish markers, delay badges, critical items, insights and an item detail drawer.
- Exports the normalised schedule as JSON.

All imported data remains local to the browser. There is no backend, authentication, database or Microsoft Project write-back.

## Local Development

```bash
npm install
npm run dev
```

## Local GitHub Tooling

Portable local copies of Git and GitHub CLI can be enabled in PowerShell with:

```powershell
.\use-github-tools.ps1
```

Then authenticate GitHub CLI:

```powershell
gh auth login --web --git-protocol https
```

## Build

```bash
npm run build
```

The production build is generated in `dist`.

## Netlify

This repo includes `netlify.toml`. Netlify can build with:

- Build command: `npm run build`
- Publish directory: `dist`

## First Test File

Use the Microsoft Project XML export supplied with the project brief:

`DAF - Project Plan V02 (1).xml`
