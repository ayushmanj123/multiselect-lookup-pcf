export interface LookupRecord {
    id: string;
    name: string;
    entityType: string;
}

export function normalizeGuid(value: string): string {
    return value.replace(/[{}]/g, "").toLowerCase();
}

export function readDatasetRecords(dataset: ComponentFramework.PropertyTypes.DataSet): LookupRecord[] {
    const fallbackType = dataset.getTargetEntityType();
    const primaryColumn = dataset.columns.find((column) => column.isPrimary)?.name;
    const records: LookupRecord[] = [];

    for (const recordId of dataset.sortedRecordIds) {
        const row = dataset.records[recordId];
        if (!row) {
            continue;
        }

        const named = row.getNamedReference();
        const id = normalizeGuid(named.id.guid || row.getRecordId());
        if (!id) {
            continue;
        }

        const namedName = named.name.trim();
        const formatted = primaryColumn ? row.getFormattedValue(primaryColumn).trim() : "";
        records.push({
            id,
            name: namedName || formatted || "(No name)",
            entityType: named.etn ?? fallbackType,
        });
    }

    return records;
}

export function mergeRecords(current: LookupRecord[], additions: LookupRecord[], removedIds: readonly string[]): LookupRecord[] {
    const removed = new Set(removedIds);
    const seen = new Set<string>();
    const merged: LookupRecord[] = [];

    for (const record of [...current, ...additions]) {
        if (!record.id || removed.has(record.id) || seen.has(record.id)) {
            continue;
        }
        seen.add(record.id);
        merged.push(record);
    }

    return merged;
}
