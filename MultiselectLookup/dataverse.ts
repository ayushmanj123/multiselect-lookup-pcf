import { IInputs } from "./generated/ManifestTypes";
import { LookupRecord, normalizeGuid } from "./records";

const EMPTY_GUID = "00000000-0000-0000-0000-000000000000";

export interface ParentRecord {
    entityType: string;
    id: string;
}

export type ParentState =
    | { kind: "ready"; record: ParentRecord }
    | { kind: "unsaved" }
    | { kind: "unavailable" };

interface TableMetadata {
    primaryIdAttribute: string;
    primaryNameAttribute: string;
    entitySetName: string;
}

interface CrmReference {
    entityType: string;
    id: string;
}

interface ActionMetadata {
    boundParameter: null;
    parameterTypes: Record<string, { typeName: string; structuralProperty: number }>;
    operationType: number;
    operationName: string;
}

interface AssociateRequest {
    target: CrmReference;
    relatedEntities: CrmReference[];
    relationship: string;
    getMetadata: () => ActionMetadata;
}

interface DisassociateRequest {
    target: CrmReference;
    relatedEntityId: string;
    relationship: string;
    getMetadata: () => ActionMetadata;
}

interface WebApiExecute {
    execute(request: AssociateRequest | DisassociateRequest): Promise<unknown>;
}

const metadataCache = new Map<string, TableMetadata>();
const viewCache = new Map<string, string>();

const entityParameter = { typeName: "mscrm.crmbaseentity", structuralProperty: 5 };
const entityCollectionParameter = { typeName: "mscrm.crmbaseentity", structuralProperty: 4 };
const stringParameter = { typeName: "Edm.String", structuralProperty: 1 };

export function readParentState(context: ComponentFramework.Context<IInputs>): ParentState {
    const mode = context.mode as ComponentFramework.Mode & {
        contextInfo?: {
            entityTypeName?: string;
            entityId?: string;
        };
    };
    const info = mode.contextInfo;
    if (!info) {
        return { kind: "unavailable" };
    }

    const entityType = typeof info.entityTypeName === "string" ? info.entityTypeName.trim() : "";
    const id = typeof info.entityId === "string" ? normalizeGuid(info.entityId) : "";
    if (!id || id === EMPTY_GUID) {
        return entityType ? { kind: "unsaved" } : { kind: "unavailable" };
    }
    if (!entityType) {
        return { kind: "unavailable" };
    }

    return { kind: "ready", record: { entityType, id } };
}

export function errorMessage(error: unknown): string {
    const message = readErrorText(error);
    return message || "The related records could not be updated.";
}

export function userFacingError(error: unknown): string {
    const message = errorMessage(error);
    if (message.toLowerCase().includes("method is not supported")) {
        return "Search and related-record changes are available on a model-driven form.";
    }
    return message;
}

export function isDuplicate(error: unknown): boolean {
    const message = readErrorText(error).toLowerCase();
    return message.includes("duplicate") || message.includes("already exists");
}

export function isDismissed(error: unknown): boolean {
    if (error === undefined || error === null || error === "") {
        return true;
    }
    const message = readErrorText(error).toLowerCase();
    if (!message) {
        return true;
    }
    return message.includes("cancel") || message.includes("dismiss") || message.includes("closed");
}

export async function searchRecords(
    context: ComponentFramework.Context<IInputs>,
    entityType: string,
    viewId: string,
    query: string,
): Promise<LookupRecord[]> {
    const table = await getTableMetadata(context, entityType);
    const fetchXml = await viewFetchXml(context, viewId, table, query);
    const response = fetchXml
        ? await context.webAPI.retrieveMultipleRecords(entityType, `?fetchXml=${encodeURIComponent(fetchXml)}`)
        : await context.webAPI.retrieveMultipleRecords(entityType, odataQuery(table, query));

    return response.entities.flatMap((entity) => {
        const record = toLookupRecord(entity, entityType, table);
        return record ? [record] : [];
    });
}

export async function openLookupDialog(
    context: ComponentFramework.Context<IInputs>,
    entityType: string,
    viewId: string,
    searchText: string,
): Promise<LookupRecord[]> {
    const options: ComponentFramework.UtilityApi.LookupOptions & { searchText?: string } = {
        allowMultiSelect: true,
        defaultEntityType: entityType,
        entityTypes: [entityType],
        searchText: searchText.trim() || undefined,
    };
    const view = normalizeGuid(viewId);
    if (view) {
        options.defaultViewId = view;
    }

    const selected = await context.utils.lookupObjects(options);
    return selected
        .filter((item) => item.id.trim().length > 0 && item.entityType.trim().length > 0)
        .map((item) => ({
            id: normalizeGuid(item.id),
            name: displayName(item.name),
            entityType: item.entityType,
        }));
}

export async function associateRecords(
    context: ComponentFramework.Context<IInputs>,
    parent: ParentRecord,
    relationshipName: string,
    related: LookupRecord[],
): Promise<void> {
    if (related.length === 0) {
        return;
    }

    const request: AssociateRequest = {
        target: { entityType: parent.entityType, id: parent.id },
        relatedEntities: related.map((record) => ({ entityType: record.entityType, id: record.id })),
        relationship: relationshipName,
        getMetadata: () => ({
            boundParameter: null,
            parameterTypes: {
                target: entityParameter,
                relatedEntities: entityCollectionParameter,
                relationship: stringParameter,
            },
            operationType: 2,
            operationName: "Associate",
        }),
    };

    const execute = readExecute(context.webAPI);
    if (execute) {
        await assertOk(await execute(request));
        return;
    }

    await postUnbound(context, "Associate", {
        target: await entityPayload(context, request.target),
        relationship: relationshipName,
        relatedEntities: await Promise.all(request.relatedEntities.map((item) => entityPayload(context, item))),
    });
}

export async function disassociateRecord(
    context: ComponentFramework.Context<IInputs>,
    parent: ParentRecord,
    relationshipName: string,
    related: LookupRecord,
): Promise<void> {
    const request: DisassociateRequest = {
        target: { entityType: parent.entityType, id: parent.id },
        relatedEntityId: related.id,
        relationship: relationshipName,
        getMetadata: () => ({
            boundParameter: null,
            parameterTypes: {
                target: entityParameter,
                relatedEntityId: stringParameter,
                relationship: stringParameter,
            },
            operationType: 2,
            operationName: "Disassociate",
        }),
    };

    const execute = readExecute(context.webAPI);
    if (execute) {
        await assertOk(await execute(request));
        return;
    }

    await postUnbound(context, "Disassociate", {
        target: await entityPayload(context, request.target),
        relationship: relationshipName,
        relatedEntity: await entityPayload(context, { entityType: related.entityType, id: related.id }),
    });
}

async function getTableMetadata(context: ComponentFramework.Context<IInputs>, logicalName: string): Promise<TableMetadata> {
    const cached = metadataCache.get(logicalName);
    if (cached) {
        return cached;
    }

    const metadata = await context.utils.getEntityMetadata(logicalName);
    const table: TableMetadata = {
        primaryIdAttribute: readString(metadata, "PrimaryIdAttribute") ?? readString(metadata, "primaryIdAttribute") ?? `${logicalName}id`,
        primaryNameAttribute: readString(metadata, "PrimaryNameAttribute") ?? readString(metadata, "primaryNameAttribute") ?? "name",
        entitySetName: readString(metadata, "EntitySetName") ?? readString(metadata, "entitySetName") ?? "",
    };
    metadataCache.set(logicalName, table);
    return table;
}

async function viewFetchXml(
    context: ComponentFramework.Context<IInputs>,
    viewId: string,
    table: TableMetadata,
    query: string,
): Promise<string | undefined> {
    const id = normalizeGuid(viewId);
    if (!id) {
        return undefined;
    }

    let fetchXml = viewCache.get(id);
    if (!fetchXml) {
        fetchXml = (await retrieveFetchXml(context, "savedquery", id)) ?? (await retrieveFetchXml(context, "userquery", id));
        if (!fetchXml) {
            return undefined;
        }
        viewCache.set(id, fetchXml);
    }

    return applySearch(fetchXml, table, query);
}

async function retrieveFetchXml(
    context: ComponentFramework.Context<IInputs>,
    tableName: "savedquery" | "userquery",
    id: string,
): Promise<string | undefined> {
    try {
        const view = await context.webAPI.retrieveRecord(tableName, id, "?$select=fetchxml");
        return readString(view, "fetchxml");
    } catch (error: unknown) {
        void error;
        return undefined;
    }
}

function applySearch(fetchXml: string, table: TableMetadata, query: string): string {
    const parsed = new DOMParser().parseFromString(fetchXml, "text/xml");
    const fetch = parsed.querySelector("fetch");
    const entity = parsed.querySelector("entity");
    if (!fetch || !entity || parsed.querySelector("parsererror")) {
        return fetchXml;
    }

    fetch.setAttribute("page", "1");
    fetch.setAttribute("count", "25");
    ensureAttribute(parsed, entity, table.primaryIdAttribute);
    ensureAttribute(parsed, entity, table.primaryNameAttribute);

    const text = query.trim();
    if (text) {
        const filter = parsed.createElement("filter");
        filter.setAttribute("type", "and");
        const condition = parsed.createElement("condition");
        condition.setAttribute("attribute", table.primaryNameAttribute);
        condition.setAttribute("operator", "like");
        condition.setAttribute("value", `%${text.replace(/%/g, "[%]").replace(/_/g, "[_]")}%`);
        filter.appendChild(condition);
        entity.appendChild(filter);
    }

    return new XMLSerializer().serializeToString(parsed);
}

function ensureAttribute(documentNode: XMLDocument, entity: Element, name: string): void {
    const children = entity.childNodes;
    for (let index = 0; index < children.length; index += 1) {
        const child = children.item(index);
        if (!child || child.nodeType !== 1 || child.nodeName !== "attribute") {
            continue;
        }
        if ((child as Element).getAttribute("name") === name) {
            return;
        }
    }
    const attribute = documentNode.createElement("attribute");
    attribute.setAttribute("name", name);
    entity.appendChild(attribute);
}

function odataQuery(table: TableMetadata, query: string): string {
    const select = `${table.primaryIdAttribute},${table.primaryNameAttribute}`;
    const text = query.trim().replace(/'/g, "''");
    const filter = text ? `&$filter=contains(${table.primaryNameAttribute},'${text}')` : "";
    return `?$select=${select}${filter}&$orderby=${table.primaryNameAttribute} asc&$top=25`;
}

function toLookupRecord(
    entity: ComponentFramework.WebApi.Entity,
    entityType: string,
    table: TableMetadata,
): LookupRecord | undefined {
    const id = readString(entity, table.primaryIdAttribute);
    if (!id) {
        return undefined;
    }
    const formatted = readString(entity, `${table.primaryNameAttribute}@OData.Community.Display.V1.FormattedValue`);
    return {
        id: normalizeGuid(id),
        name: readString(entity, table.primaryNameAttribute) ?? formatted ?? "(No name)",
        entityType,
    };
}

function readExecute(webAPI: ComponentFramework.WebApi): WebApiExecute["execute"] | undefined {
    if (!("execute" in webAPI)) {
        return undefined;
    }
    const candidate = webAPI as WebApiExecute;
    return typeof candidate.execute === "function" ? candidate.execute.bind(candidate) : undefined;
}

async function entityPayload(
    context: ComponentFramework.Context<IInputs>,
    reference: CrmReference,
): Promise<Record<string, string>> {
    const table = await getTableMetadata(context, reference.entityType);
    return {
        "@odata.type": `Microsoft.Dynamics.CRM.${reference.entityType}`,
        [table.primaryIdAttribute]: normalizeGuid(reference.id),
    };
}

async function postUnbound(
    context: ComponentFramework.Context<IInputs>,
    action: "Associate" | "Disassociate",
    body: unknown,
): Promise<void> {
    const response = await fetch(`${clientUrl(context)}/api/data/v9.2/${action}`, {
        method: "POST",
        credentials: "same-origin",
        headers: {
            Accept: "application/json",
            "Content-Type": "application/json; charset=utf-8",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
        },
        body: JSON.stringify(body),
    });
    if (!response.ok) {
        throw new Error(await responseError(response));
    }
}

async function assertOk(result: unknown): Promise<void> {
    if (result instanceof Response && !result.ok) {
        throw new Error(await responseError(result));
    }
}

async function responseError(response: Response): Promise<string> {
    const text = await response.text();
    if (!text.trim()) {
        return response.statusText || "The request failed.";
    }
    try {
        return errorMessage(parseUnknown(text));
    } catch (error: unknown) {
        void error;
        return text;
    }
}

function clientUrl(context: ComponentFramework.Context<IInputs>): string {
    const page = (context as { page?: { getClientUrl?: () => string } }).page;
    const url = page?.getClientUrl?.() ?? window.location.origin;
    return url.replace(/\/$/, "");
}

function displayName(name: string | undefined): string {
    const trimmed = name?.trim() ?? "";
    return trimmed.length > 0 ? trimmed : "(No name)";
}

function readString(source: object, key: string): string | undefined {
    const value: unknown = Reflect.get(source, key) as unknown;
    return typeof value === "string" && value.trim().length > 0 ? value : undefined;
}

function readErrorText(error: unknown): string {
    if (error instanceof Error) {
        return error.message.trim();
    }
    if (typeof error === "string") {
        return error.trim();
    }
    if (!isRecord(error)) {
        return "";
    }
    if (typeof error.message === "string") {
        return error.message.trim();
    }
    if (isRecord(error.error) && typeof error.error.message === "string") {
        return error.error.message.trim();
    }
    return "";
}

function isRecord(value: unknown): value is Record<string, unknown> {
    return typeof value === "object" && value !== null;
}

function parseUnknown(text: string): unknown {
    return JSON.parse(text) as unknown;
}
