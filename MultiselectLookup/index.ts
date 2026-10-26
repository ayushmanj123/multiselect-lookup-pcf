import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { readParentState } from "./dataverse";
import { LookupControl } from "./LookupControl";
import * as React from "react";

export class MultiselectLookup implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private pagingReady = false;
    private pagesLoaded = 0;
    private formKey = "";

    public init(context: ComponentFramework.Context<IInputs>): void {
        context.mode.trackContainerResize(true);
    }

    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        this.loadRemainingRecords(context);
        return React.createElement(LookupControl, { context });
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {
        this.pagingReady = false;
    }

    private loadRemainingRecords(context: ComponentFramework.Context<IInputs>): void {
        const dataset = context.parameters.relatedRecords;
        const formRecord = readParentState(context);
        const formKey = formRecord.kind === "ready" ? formRecord.record.id : formRecord.kind;
        if (formKey !== this.formKey) {
            this.formKey = formKey;
            this.pagesLoaded = 0;
            this.pagingReady = false;
        }
        if (dataset.loading) {
            return;
        }
        if (!this.pagingReady) {
            if (dataset.paging.pageSize <= 0) {
                return;
            }
            this.pagingReady = true;
            if (dataset.paging.pageSize < 250) {
                try {
                    dataset.paging.setPageSize(250);
                    dataset.refresh();
                    return;
                } catch (error: unknown) {
                    void error;
                }
            }
        }
        if (dataset.paging.hasNextPage && this.pagesLoaded < 40) {
            this.pagesLoaded += 1;
            dataset.paging.loadNextPage();
        }
    }
}
