# Project status and engineering notes

This document records the boundary between what this repository is intended to demonstrate, what has been verified in the repository, and what still needs evidence. It is deliberately more specific than a feature list.

## Current stage

**Windows desktop conversion utility**

## Why this exists

I built this to turn a repeated Office-to-PDF conversion task into a reviewable batch process with predictable output locations.

## Scope and known limitations

It is tested for Windows desktop use and depends on locally installed Microsoft Office components. Conversion fidelity, protected files, macros, fonts, and damaged documents can produce failures or different PDFs. It is not a headless server converter.

## Next evidence to collect

Record supported Office versions and add representative fixtures for layout drift, password protection, and missing-font failures.

## Maintenance rule

Future changes should describe one concrete behavior, include the smallest relevant verification step, and update this document whenever the project boundary changes.
