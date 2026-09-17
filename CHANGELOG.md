# Changelog

## [1.0.0] - 2026-09-17

### Added
- Public sanitized Python baseline.
- External `.env` configuration and safe `.env.example`.
- Inventory cleaning and reconciliation.
- Canonical tag normalization.
- FOCUS CSV valuation.
- Configurable exclusions and orphan reporting.
- Shared-cost distribution in application summaries.
- Tagged-cost coverage KPI.
- Excel reporting and audit outputs.
- Public README and detailed documentation.

### Sanitized
- Removed client/organization identifiers from public code/configuration.
- Removed personal filesystem paths.
- Removed organization-specific application/environment mappings.
- Excluded operational datasets and generated reports from source control.

### Notes
This release establishes the documented public/certification baseline and intentionally retains the single-script implementation to minimize functional change from the validated operational artifact.
