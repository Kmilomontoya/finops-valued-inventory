# FOCUS Valued Resource Inventory

Reusable Python FinOps artifact that enriches a cloud resource inventory with monthly cost and usage data in **FOCUS** format.

## FinOps purpose

The artifact answers three operational questions:

1. What did each inventoried resource cost during the period?
2. What percentage of inventory cost has valid allocation metadata?
3. Where is unallocated or insufficiently tagged cost concentrated?

The operating documentation frames the implementation around **Allocation** and **Reporting & Analytics**, with supporting ingestion, normalization, reconciliation, and data-quality controls.

## Processing flow

```text
Resource inventory
       |
       v
L1 - Cleaning
       |
       v
L2 - Canonical tag normalization
       |
       +------------------+
       |                  |
       v                  |
FOCUS cost & usage -------+
       |
       v
L3 - Valuation / reconciliation
       |
       +--> Orphan FOCUS cost
       +--> Excluded records
       +--> Audit controls
       |
       v
Valued inventory + summaries + tagged-cost KPI
```

The inventory is the dominant dataset. FOCUS is used to attach cost to that population. Unmatched FOCUS records are reported separately rather than silently discarded.

## Requirements

- Python 3.10+
- pandas
- openpyxl
- python-dotenv

```bash
pip install -r requirements.txt
```

## Quick start

1. Copy `.env.example` to `.env`.
2. Configure local input/output paths in `.env`.
3. Place the resource inventory (`.xlsx`) and FOCUS CSV file(s) in the local input directory.
4. Run:

```bash
python inventario_valorizado.py
```

The interactive menu verifies required inputs before processing.

> Never commit `.env`, real inventories, FOCUS exports, generated reports, resource IDs, subscription identifiers, owner information, or real cost data.

## Inputs

The preferred inventory reconciliation key is `ResourceId`. If unavailable, the tool can fall back to resource name + resource group. One or more FOCUS CSV files can be processed. The valuation cost field is controlled by `COST_COLUMN`.

See [Configuration](docs/configuration.md).

## Processing stages

**L1 - Cleaning:** detects the inventory header, maps supported base aliases, cleans values, establishes the join key, removes duplicate join keys, and derives subscription information from ResourceId when applicable.

**L2 - Normalization:** consolidates configured tag variants into canonical fields, applies configured environment/application mappings, preserves supported shared-cost expressions, and records conflicts/unmapped values for audit.

**L3 - Valuation:** reads FOCUS CSVs, applies configured exclusions, aggregates cost by join key, maps cost to the inventory, and identifies unmatched FOCUS cost.

See [Architecture](docs/architecture.md) and [Methodology](docs/methodology.md).

## Outputs

The valued-inventory workbook can contain `Inventario_Valorizado`, `Resumen_APPLICATION`, `Resumen_ENVIRONMENT`, `Resumen_TYPE`, `Huerfanos_FOCUS`, `Excluidos`, and `Auditoria`.

The KPI workbook can contain `Resumen`, `KPI_Cobertura`, `Gap_APPLICATION`, and `Cobertura_por_RG`. Exact sheets depend on available/configured fields.

## Tagged-cost coverage

```text
Tagged Cost Coverage =
Cost of resources with a valid tag
---------------------------------- x 100
Total valued inventory cost
```

Placeholder values such as `NO_DEFINIDO` and `NO_ETIQUETADO` do not count as valid coverage. Resource-count coverage is supporting context; the primary KPI is cost-weighted.

## Data protection

This public version is intentionally sanitized. It contains no client datasets, real resource identifiers, subscriptions, owner information, credentials, or organization-specific application mappings. Operational data and generated outputs remain outside source control.

## Limitations

- `ResourceId` matching is preferred; name + resource-group fallback can be ambiguous across subscriptions.
- Inventory and FOCUS data should represent compatible populations/periods.
- Canonical mappings are configuration-driven and must be adapted locally.
- The tool performs analysis/reporting only; it does not alter cloud resources or tags.

## Documentation

- [Methodology](docs/methodology.md)
- [Architecture](docs/architecture.md)
- [Configuration](docs/configuration.md)
- [Data model](docs/data_model.md)
- [Troubleshooting](docs/troubleshooting.md)
- [Changelog](CHANGELOG.md)
- [Notice](NOTICE.md)

## Version

**1.0.0 - public sanitized documentation baseline**

No open-source license is asserted in this package. See [NOTICE.md](NOTICE.md).
