---
name: mbg-penerima-import
description: Generate per-group Excel files from a Profil Mitra SPPG workbook using the penerima import template, with unique 12-digit NISN values and wajar birthdates.
---

# MBG Penerima Import

Use this skill when the user wants to turn a `Profil Mitra SPPG.xlsx` workbook into one or more `template_import_penerima.xlsx`-style output files.

## Workflow

1. Read the profile workbook and locate the active sheet.
2. Detect the source columns from the header row, then read:
   - `Jenis Kelompok`
   - `Nama Kelompok`
   - `Jumlah Pria`
   - `Jumlah Wanita`
   - `Jumlah Guru/Kader`
   - `Jumlah Tendik`
3. Read the penerima template and keep rows 1 and 2 unchanged.
4. Remove sample data from row 3 onward before writing new rows.
5. Generate one output workbook per source row unless the user explicitly asks for aggregation.
6. Fill data starting at row 3.

## Output Rules

- `NISN/NIK`: always generate unique 12-digit numbers for all rows.
- `Jenis Kelamin ID`: `392` for male, `64` for female.
- `Posisi`: `1` for peserta didik, `2` for guru/kader, `3` for tendik.
- `Status Menerima`: always `true` unless the user asks otherwise.
- Name values should be Indonesian and plausible for the category.
- Date of birth should be plausible for the category:
  - bayi/balita/PAUD/TK: early childhood
  - SD kelas 1-3: child age
  - SD kelas 4-6: older child age
  - SMP/SMA/SMK/MA: teenager age
  - guru/kader and tendik: adult age

## File Naming

Use a safe filename from `Jenis Kelompok` plus `Nama Kelompok`, removing invalid filename characters.

## Validation

Before finishing, verify:

- Header rows remain unchanged.
- All generated NISN values are unique.
- The count of written rows matches the requested totals.
- All date values are valid Excel dates in `DD-MM-YYYY` format.

## Tooling

Prefer the bundled generator script in `scripts/generate_penerima.py` for deterministic output.
