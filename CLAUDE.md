# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Google Apps Script project for MOOMOOZ ESSENTIAL (무무즈에센셜) — a children's clothing brand. It's a one-time setup script that creates a production management system in Google Drive/Sheets/Docs.

## Development Workflow (CLASP)

This project uses [CLASP](https://github.com/google/clasp) to sync Apps Script code locally.

```bash
# Push local changes to Apps Script
clasp push

# Pull latest from Apps Script
clasp pull

# Open in browser editor
clasp open

# Deploy a new version
clasp deploy
```

Script ID is in `.clasp.json`. Runtime target is Apps Script V8 (modern JS). All source files use `.js` or `.gs` extensions.

## Architecture

All logic lives in a single file: `Code.js` (~518 lines).

### Entry Point

`setup()` — one-time initialization function. Run this once in the Apps Script editor to bootstrap the entire system. It sequentially:
1. Creates a folder hierarchy under a Google Shared Drive
2. Generates a Master DB spreadsheet (6 sheets)
3. Generates an Operations spreadsheet (4 sheets)
4. Generates 5 Work Order Doc templates (one per garment type)

### Key Functions

| Function | Purpose |
|---|---|
| `setup()` | Main entry point — runs everything once |
| `createFolderStructure()` | Creates 5 named folders in the shared drive |
| `createMasterDB()` | Builds the Master DB spreadsheet with all reference sheets |
| `createOperationFile()` | Builds the Operations spreadsheet for order tracking |
| `createWorkOrderTemplates()` | Generates Google Docs templates for each item type |
| `setHeader()` | Applies styled header rows with optional data validation dropdowns |
| `getConfig()` / `getMasterSSId()` | Reads system config from the System Config sheet |

### Data Model

- **상품마스터** (Product Master): product codes, names, categories (BABY/KIDS), item types (TOP/BOTTOM/OUTER/SET/DRESS)
- **BOM**: material requirements per product with loss rates
- **자재마스터** (Material Master): fabric/material inventory
- **업체마스터** (Vendor Master): supplier contacts
- **사이즈스펙** (Size Specs): size grading per item type
- **작업지시서목록** (Work Order List): tracks all issued work orders

### Configuration / Secrets

Sensitive values (Drive IDs, etc.) are stored as **Script Properties**, not in source code.

- Runtime: `PropertiesService.getScriptProperties().getProperty('KEY')`
- Set values in: Apps Script 에디터 > 프로젝트 설정 > 스크립트 속성
- Local reference: `.env` (gitignored) — copy from `.env.example`

| Property key | Description |
|---|---|
| `SHARED_DRIVE_ID` | Google Shared Drive root folder ID |

After `setup()` runs, all generated file/folder IDs are also written to the 시스템설정 sheet and readable via `getConfig('KEY')`.

## No Test Infrastructure

There are no automated tests. Manual testing is done by running `setup()` in the Apps Script editor and verifying the generated Drive/Sheets/Docs output.
