import { FluentProvider, webDarkTheme, webLightTheme, type Theme } from "@fluentui/react-components";
import * as React from "react";
import { IInputs } from "./generated/ManifestTypes";
import {
    associateRecords,
    disassociateRecord,
    isDismissed,
    isDuplicate,
    openLookupDialog,
    readParentState,
    searchRecords,
    userFacingError,
} from "./dataverse";
import { LookupRecord, mergeRecords, readDatasetRecords } from "./records";

export interface LookupControlProps {
    context: ComponentFramework.Context<IInputs>;
}

interface LookupStyle extends React.CSSProperties {
    "--msl-font": string;
    "--msl-font-size": string;
    "--msl-line-height": string;
    "--msl-text": string;
    "--msl-text-disabled": string;
    "--msl-placeholder": string;
    "--msl-link": string;
    "--msl-icon": string;
    "--msl-bg": string;
    "--msl-bg-disabled": string;
    "--msl-stroke": string;
    "--msl-stroke-disabled": string;
    "--msl-focus": string;
    "--msl-hover": string;
    "--msl-pill": string;
    "--msl-danger": string;
    "--msl-shadow": string;
}

let controlSequence = 0;

export const LookupControl: React.FunctionComponent<LookupControlProps> = ({ context }) => {
    const dataset = context.parameters.relatedRecords;
    const entityType = dataset.getTargetEntityType();
    const viewId = dataset.getViewId();
    const relationshipName = context.parameters.relationshipName.raw?.trim() ?? "";
    const formRecord = readParentState(context);
    const disabled = context.mode.isControlDisabled || formRecord.kind === "unsaved";
    const serverRecords = readDatasetRecords(dataset);
    const serverKey = serverRecords.map((record) => record.id).join("|");

    const [query, setQuery] = React.useState("");
    const [open, setOpen] = React.useState(false);
    const [activeIndex, setActiveIndex] = React.useState(0);
    const [results, setResults] = React.useState<LookupRecord[]>([]);
    const [searching, setSearching] = React.useState(false);
    const [searchError, setSearchError] = React.useState("");
    const [actionError, setActionError] = React.useState("");
    const [pendingAdds, setPendingAdds] = React.useState<LookupRecord[]>([]);
    const [pendingRemoves, setPendingRemoves] = React.useState<string[]>([]);
    const [busy, setBusy] = React.useState(false);

    const contextRef = React.useRef(context);
    contextRef.current = context;
    const inputRef = React.useRef<HTMLInputElement>(null);
    const activeRef = React.useRef<HTMLButtonElement | null>(null);
    const controlId = React.useRef("");
    if (!controlId.current) {
        controlSequence += 1;
        controlId.current = `msl-${controlSequence}`;
    }
    const listId = `${controlId.current}-list`;

    const selected = mergeRecords(serverRecords, pendingAdds, pendingRemoves);
    const selectedIds = new Set(selected.map((record) => record.id));
    const optionCount = results.length + 1;
    const safeIndex = optionCount === 0 ? 0 : ((activeIndex % optionCount) + optionCount) % optionCount;
    const label = context.mode.label.trim() || "Lookup";
    const placeholder =
        formRecord.kind === "unsaved"
            ? "Save the record to select related records"
            : selected.length === 0
              ? "Look for records"
              : "";
    const notice = actionError || (dataset.error ? dataset.errorMessage : "");

    React.useEffect(() => {
        const serverIds = new Set(serverKey.split("|").filter((id) => id.length > 0));
        setPendingAdds((current) => current.filter((record) => !serverIds.has(record.id)));
        setPendingRemoves((current) => current.filter((id) => serverIds.has(id)));
    }, [serverKey]);

    React.useEffect(() => {
        if (!open || disabled || !entityType) {
            return undefined;
        }

        let cancelled = false;
        const timer = window.setTimeout(() => {
            setSearching(true);
            void searchRecords(contextRef.current, entityType, viewId, query)
                .then((records) => {
                    if (cancelled) {
                        return undefined;
                    }
                    setResults(records);
                    setSearchError("");
                    setActiveIndex(0);
                    return undefined;
                })
                .catch((error: unknown) => {
                    if (cancelled) {
                        return undefined;
                    }
                    setResults([]);
                    setSearchError(userFacingError(error));
                    return undefined;
                })
                .finally(() => {
                    if (!cancelled) {
                        setSearching(false);
                    }
                });
        }, 250);

        return () => {
            cancelled = true;
            window.clearTimeout(timer);
        };
    }, [disabled, entityType, open, query, viewId]);

    React.useEffect(() => {
        activeRef.current?.scrollIntoView({ block: "nearest" });
    }, [open, safeIndex]);

    const requireLink = (): { relationship: string; parentId: string; parentType: string } | undefined => {
        if (!relationshipName) {
            setActionError("Set Relationship Name to the schema name of the subgrid relationship.");
            return undefined;
        }
        if (formRecord.kind === "unsaved") {
            setActionError("Save the record before selecting related records.");
            return undefined;
        }
        if (formRecord.kind !== "ready") {
            setActionError("Related records can be changed on a model-driven form.");
            return undefined;
        }
        return { relationship: relationshipName, parentId: formRecord.record.id, parentType: formRecord.record.entityType };
    };

    const addRecords = (records: LookupRecord[]): void => {
        const link = requireLink();
        if (!link || busy) {
            return;
        }
        const additions = records.filter((record) => record.id && !selectedIds.has(record.id));
        if (additions.length === 0) {
            return;
        }

        setPendingAdds((current) => mergeRecords(current, additions, []));
        setBusy(true);
        setActionError("");
        void associateRecords(
            contextRef.current,
            { entityType: link.parentType, id: link.parentId },
            link.relationship,
            additions,
        )
            .then(() => {
                contextRef.current.parameters.relatedRecords.refresh();
                return undefined;
            })
            .catch((error: unknown) => {
                if (isDuplicate(error)) {
                    contextRef.current.parameters.relatedRecords.refresh();
                } else {
                    const addedIds = new Set(additions.map((record) => record.id));
                    setPendingAdds((current) => current.filter((record) => !addedIds.has(record.id)));
                    setActionError(userFacingError(error));
                }
                return undefined;
            })
            .finally(() => {
                setBusy(false);
            });
    };

    const removeRecord = (record: LookupRecord): void => {
        const link = requireLink();
        if (!link || busy) {
            return;
        }

        setPendingAdds((current) => current.filter((item) => item.id !== record.id));
        setPendingRemoves((current) => (current.includes(record.id) ? current : [...current, record.id]));
        setBusy(true);
        setActionError("");
        void disassociateRecord(
            contextRef.current,
            { entityType: link.parentType, id: link.parentId },
            link.relationship,
            record,
        )
            .then(() => {
                contextRef.current.parameters.relatedRecords.refresh();
                return undefined;
            })
            .catch((error: unknown) => {
                setPendingRemoves((current) => current.filter((id) => id !== record.id));
                    setActionError(userFacingError(error));
                return undefined;
            })
            .finally(() => {
                setBusy(false);
                inputRef.current?.focus();
            });
    };

    const openRecord = (record: LookupRecord): void => {
        dataset.openDatasetItem({
            id: { guid: record.id },
            name: record.name,
            etn: record.entityType,
        });
    };

    const openDialog = (): void => {
        if (disabled || !entityType || busy) {
            return;
        }
        setOpen(false);
        void openLookupDialog(contextRef.current, entityType, viewId, query)
            .then((picked) => {
                addRecords(picked);
                setQuery("");
                return undefined;
            })
            .catch((error: unknown) => {
                if (!isDismissed(error)) {
                    setActionError(userFacingError(error));
                }
                return undefined;
            });
    };

    const choose = (record: LookupRecord): void => {
        if (selectedIds.has(record.id)) {
            removeRecord(record);
        } else {
            addRecords([record]);
        }
        setQuery("");
    };

    const onKeyDown = (event: React.KeyboardEvent<HTMLInputElement>): void => {
        if (event.key === "ArrowDown") {
            event.preventDefault();
            setOpen(true);
            setActiveIndex((index) => (index + 1) % optionCount);
            return;
        }
        if (event.key === "ArrowUp") {
            event.preventDefault();
            setOpen(true);
            setActiveIndex((index) => (index - 1 + optionCount) % optionCount);
            return;
        }
        if (event.key === "Enter") {
            event.preventDefault();
            if (!open) {
                setOpen(true);
                return;
            }
            if (safeIndex >= results.length) {
                openDialog();
                return;
            }
            const record = results[safeIndex];
            if (record) {
                choose(record);
            }
            return;
        }
        if (event.key === "Escape") {
            setOpen(false);
            return;
        }
        if (event.key === "Backspace" && query.length === 0) {
            const last = selected[selected.length - 1];
            if (last) {
                removeRecord(last);
            }
        }
    };

    const onBlur = (event: React.FocusEvent<HTMLDivElement>): void => {
        const next = event.relatedTarget;
        if (next instanceof Node && event.currentTarget.contains(next)) {
            return;
        }
        setOpen(false);
    };

    const theme = hostTheme(context);
    const style: LookupStyle = {
        "--msl-font": theme.fontFamilyBase,
        "--msl-font-size": theme.fontSizeBase300,
        "--msl-line-height": theme.lineHeightBase300,
        "--msl-text": theme.colorNeutralForeground1,
        "--msl-text-disabled": theme.colorNeutralForegroundDisabled,
        "--msl-placeholder": theme.colorNeutralForeground4,
        "--msl-link": theme.colorBrandForegroundLink,
        "--msl-icon": theme.colorNeutralForeground2,
        "--msl-bg": theme.colorNeutralBackground1,
        "--msl-bg-disabled": theme.colorNeutralBackgroundDisabled,
        "--msl-stroke": theme.colorNeutralStroke1,
        "--msl-stroke-disabled": theme.colorNeutralStrokeDisabled,
        "--msl-focus": theme.colorCompoundBrandStroke,
        "--msl-hover": theme.colorNeutralBackground1Hover,
        "--msl-pill": theme.colorNeutralBackground3,
        "--msl-danger": theme.colorPaletteRedForeground1,
        "--msl-shadow": theme.shadow16,
    };

    return (
        <FluentProvider theme={theme} style={{ background: "transparent", width: "100%" }}>
            <div className="msl" style={style} onBlur={onBlur}>
                <div className="msl-field" data-disabled={disabled ? "true" : "false"}>
                    <div className="msl-values">
                        {selected.map((record) => (
                            <span className="msl-pill" key={record.id}>
                                <button type="button" className="msl-pill-name" title={record.name} onClick={() => { openRecord(record); }}>
                                    {record.name}
                                </button>
                                {!disabled && (
                                    <button
                                        type="button"
                                        className="msl-pill-remove"
                                        aria-label={`Remove ${record.name}`}
                                        disabled={busy}
                                        onMouseDown={(event) => { event.preventDefault(); }}
                                        onClick={() => { removeRecord(record); }}
                                    >
                                        <DismissIcon />
                                    </button>
                                )}
                            </span>
                        ))}
                        <input
                            ref={inputRef}
                            className="msl-input"
                            role="combobox"
                            aria-label={label}
                            aria-expanded={open}
                            aria-controls={listId}
                            aria-autocomplete="list"
                            aria-activedescendant={open ? optionElementId(listId, safeIndex, results.length) : undefined}
                            aria-busy={busy || dataset.loading}
                            autoComplete="off"
                            spellCheck={false}
                            placeholder={placeholder}
                            value={query}
                            disabled={disabled}
                            onFocus={() => { setOpen(true); }}
                            onChange={(event) => {
                                setQuery(event.target.value);
                                setOpen(true);
                            }}
                            onKeyDown={onKeyDown}
                        />
                    </div>
                    <button
                        type="button"
                        className="msl-search"
                        aria-label="Search records"
                        disabled={disabled || !entityType || busy}
                        onClick={openDialog}
                    >
                        <SearchIcon />
                    </button>
                </div>
                {open && !disabled && (
                    <div className="msl-popup" id={listId} role="listbox" aria-label={label} aria-multiselectable="true">
                        {searching && results.length === 0 && (
                            <div className="msl-status" role="status">
                                <span className="msl-spinner" />
                                Searching
                            </div>
                        )}
                        {!searching && searchError && <div className="msl-status">{searchError}</div>}
                        {!searching && !searchError && results.length === 0 && (
                            <div className="msl-status">No records found</div>
                        )}
                        {results.map((record, index) => (
                            <button
                                key={record.id}
                                ref={index === safeIndex ? activeRef : undefined}
                                id={`${listId}-option-${index}`}
                                type="button"
                                role="option"
                                aria-selected={selectedIds.has(record.id)}
                                className="msl-option"
                                data-active={index === safeIndex ? "true" : "false"}
                                onMouseDown={(event) => { event.preventDefault(); }}
                                onMouseEnter={() => { setActiveIndex(index); }}
                                onClick={() => { choose(record); }}
                            >
                                <span className="msl-option-name">{record.name}</span>
                                {selectedIds.has(record.id) && (
                                    <span className="msl-check" aria-hidden="true">
                                        <CheckIcon />
                                    </span>
                                )}
                            </button>
                        ))}
                        <button
                            ref={safeIndex >= results.length ? activeRef : undefined}
                            id={`${listId}-more`}
                            type="button"
                            className="msl-more"
                            data-active={safeIndex >= results.length ? "true" : "false"}
                            onMouseDown={(event) => { event.preventDefault(); }}
                            onMouseEnter={() => { setActiveIndex(results.length); }}
                            onClick={openDialog}
                        >
                            <SearchIcon />
                            Look up more records
                        </button>
                    </div>
                )}
                {notice && <div className="msl-error" role="alert">{notice}</div>}
            </div>
        </FluentProvider>
    );
};

function optionElementId(listId: string, index: number, resultCount: number): string {
    return index >= resultCount ? `${listId}-more` : `${listId}-option-${index}`;
}

function hostTheme(context: ComponentFramework.Context<IInputs>): Theme {
    const fallback = context.fluentDesignLanguage?.isDarkTheme ? webDarkTheme : webLightTheme;
    const host: unknown = context.fluentDesignLanguage ? context.fluentDesignLanguage.tokenTheme : undefined;
    if (!isTheme(host)) {
        return fallback;
    }
    return host;
}

function isTheme(value: unknown): value is Theme {
    if (!value || typeof value !== "object") {
        return false;
    }
    const candidate = value as { colorNeutralBackground1?: unknown; fontFamilyBase?: unknown };
    return typeof candidate.colorNeutralBackground1 === "string" && typeof candidate.fontFamilyBase === "string";
}

const SearchIcon: React.FunctionComponent = () => (
    <svg width="16" height="16" viewBox="0 0 20 20" fill="currentColor" aria-hidden="true">
        <path d="M8.5 3a5.5 5.5 0 0 1 4.23 9.02l4.12 4.13a.5.5 0 0 1-.63.76l-.07-.06-4.13-4.12A5.5 5.5 0 1 1 8.5 3Zm0 1a4.5 4.5 0 1 0 0 9 4.5 4.5 0 0 0 0-9Z" />
    </svg>
);

const DismissIcon: React.FunctionComponent = () => (
    <svg width="12" height="12" viewBox="0 0 12 12" fill="currentColor" aria-hidden="true">
        <path d="M2.09 2.22a.5.5 0 0 1 .63-.06l.07.06L6 5.29l3.21-3.07a.5.5 0 0 1 .7.7L6.71 6l3.2 3.08a.5.5 0 0 1-.63.76l-.07-.06L6 6.71 2.79 9.78a.5.5 0 0 1-.7-.7L5.29 6 2.09 2.92a.5.5 0 0 1 0-.7Z" />
    </svg>
);

const CheckIcon: React.FunctionComponent = () => (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true">
        <path d="M13.78 4.22a.75.75 0 0 1 0 1.06l-6.25 6.25a.75.75 0 0 1-1.06 0L2.22 7.28a.75.75 0 1 1 1.06-1.06L7 9.94l5.72-5.72a.75.75 0 0 1 1.06 0Z" />
    </svg>
);
