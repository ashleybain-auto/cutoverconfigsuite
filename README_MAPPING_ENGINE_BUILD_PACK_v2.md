# Mapping Engine Build Pack - Revision 2

Date: 2026-08-20

This revision incorporates the `ConfigTool PowerApp.zip` extract into the original PHA/REG Excel mapping-engine assessment.

## Updated documents
- 01_Reverse_Engineering_Assessment_v2.docx
- 02_Target_Solution_Architecture_and_Data_Design_v2.docx
- 03_Build_From_Scratch_Implementation_Guide_v2.docx
- 04_Testing_Validation_Performance_Optimization_Strategy_v2.docx

## New evidence incorporated
- Exported VBA modules: ConfigTools_MedOrder.bas, Pha_Crosswalks_ConfigTools.bas, DeveloperFunctions.bas, DeveloperFunctions_MedOrder.bas, ThisWorkbook.cls, and worksheet class exports.
- Configuration/control-plane workbooks: ConnectionRegistry, EntityMapping, QueueConfig, SecurityMatrix, ThemeConfig, BulkUploadTemplate, and AuditRollback.

## Major revision points
1. Explicit environment/connection/table routing control plane.
2. Worker-aware route execution and dual PostgreSQL / SQL Server join-plan metadata.
3. Durable pause/resume queue, FIFO replay, retry classes, dry-run and failure-code handling.
4. Role, permission and field-level security model.
5. Bulk upload split into batch metadata and row payload contracts.
6. Immutable audit/rollback and connection-event concepts.
7. Metadata-driven theme/UX configuration.
8. Stronger VBA parity test cases and a reduced reverse-engineering limitation.
9. Route-aware performance, recovery and concurrency test scenarios.

## Important caveats
- Configuration defaults such as queue depth 10,000 and batch size 500 are recommendations in the supplied implementation package, not production capacity guarantees.
- The exported VBA source materially improves parity analysis, but the final production source of truth should still be reconciled to the controlled production XLSB versions and exact stored-procedure contracts.
- Production deployment should store only secret references in configuration; do not migrate live passwords/tokens from workbook/package artifacts.
