import * as React from "react";
import { IInputs } from "./generated/ManifestTypes";
import { Icon } from "@fluentui/react/lib/Icon";
import { Stack, IStackTokens, IStackStyles } from "@fluentui/react/lib/Stack";
import { TextField } from "@fluentui/react/lib/TextField";
import { ActivityItem } from "@fluentui/react/lib/ActivityItem";
import { Text } from "@fluentui/react/lib/Text";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { Separator } from "@fluentui/react/lib/Separator";
import { Callout, DirectionalHint } from "@fluentui/react/lib/Callout";
import { IconButton } from "@fluentui/react/lib/Button";
import { ContextualMenu, IContextualMenuItem, DirectionalHint as MenuDirectionalHint } from "@fluentui/react/lib/ContextualMenu";


export interface IHistoryTooltipProps {
    context: ComponentFramework.Context<IInputs>;
    value: string | null;
    fieldName: string;
    onChange: (newValue: string | undefined) => void;
}

interface IAuditRecord {
    createdon: string;
    userid: string;
    operation: string;
    action: string;
    auditId: string;
    oldValue?: string;
    newValue?: string;
}

interface IOptionMetadata {
    Options?: {
        Value: string | number;
        Label: string;
    }[];
    Type?: string;
}

const stackTokens: IStackTokens = { childrenGap: 5 };
const tooltipStackStyles: IStackStyles = {
    root: {
        padding: "8px",
        maxWidth: "300px",
        background: "rgba(255, 255, 255, 0.95)",
        backdropFilter: "blur(10px)",
        borderRadius: "8px",
        boxShadow: "0 4px 15px rgba(0,0,0,0.1)",
        border: "1px solid #edebe9"
    }
};



const TimestampToggler: React.FunctionComponent<{ dateStr: string }> = ({ dateStr }) => {
    const [timeFormat, setTimeFormat] = React.useState<'local' | 'utc'>('local');
    const [showMenu, setShowMenu] = React.useState(false);
    const linkRef = React.useRef<HTMLSpanElement>(null);

    const date = new Date(dateStr);
    const localString = date.toLocaleString(); // e.g. "10/2/2026, 4:03:22 PM"

    const amPmRegex = /^(.*)\s(AM|PM)$/i;
    const match = localString.match(amPmRegex);

    const onMenuDismiss = () => setShowMenu(false);

    const menuStyles = {
        root: { minWidth: 120 },
        subComponentStyles: {
            menuItem: {
                root: { height: 32, minHeight: 32, lineHeight: 32 },
                linkContent: { height: 32, lineHeight: 32, padding: "0 10px" },
                icon: { fontSize: 12, lineHeight: 32, padding: 0 },
                label: { fontSize: 13, lineHeight: 32, margin: 0 }
            }
        }
    };

    const menuItems: IContextualMenuItem[] = [
        {
            key: 'local',
            text: 'Local Time',
            iconProps: { iconName: 'Clock' },
            onClick: () => setTimeFormat('local'),
            checked: timeFormat === 'local',
            canCheck: true,
        },
        {
            key: 'utc',
            text: 'UTC Time',
            iconProps: { iconName: 'World' },
            onClick: () => setTimeFormat('utc'),
            checked: timeFormat === 'utc',
            canCheck: true,
        }
    ];

    if (timeFormat === 'utc') {
        const utcTime = `${date.getUTCHours().toString().padStart(2, '0')}:${date.getUTCMinutes().toString().padStart(2, '0')}:${date.getUTCSeconds().toString().padStart(2, '0')} UTC`;
        const localDatePart = date.toLocaleDateString();

        return (
            <span>
                {localDatePart}, {utcTime.replace(" UTC", "")}{" "}
                <span
                    ref={linkRef}
                    style={{ cursor: "pointer", color: "#0078d4", fontWeight: 600, textDecoration: "underline", display: "inline-flex", alignItems: "center" }}
                    onClick={(e) => { e.stopPropagation(); setShowMenu(!showMenu); }}
                >
                    UTC
                    <Icon iconName="ChevronDown" style={{ fontSize: 10, marginLeft: 2 }} />
                </span>
                {showMenu && (
                    <ContextualMenu
                        items={menuItems}
                        target={linkRef}
                        onDismiss={onMenuDismiss}
                        directionalHint={MenuDirectionalHint.bottomLeftEdge}
                        styles={menuStyles}
                    />
                )}
            </span>
        );
    }

    if (match) {
        // match[1] is "10/2/2026, 4:03:22"
        // match[2] is "PM"
        return (
            <span>
                {match[1]}{" "}
                <span
                    ref={linkRef}
                    style={{ cursor: "pointer", color: "#0078d4", fontWeight: 600, textDecoration: "underline", display: "inline-flex", alignItems: "center" }}
                    onClick={(e) => { e.stopPropagation(); setShowMenu(!showMenu); }}
                >
                    {match[2].toUpperCase()}
                    <Icon iconName="ChevronDown" style={{ fontSize: 10, marginLeft: 2 }} />
                </span>
                {showMenu && (
                    <ContextualMenu
                        items={menuItems}
                        target={linkRef}
                        onDismiss={onMenuDismiss}
                        directionalHint={MenuDirectionalHint.bottomLeftEdge}
                        styles={menuStyles}
                    />
                )}
            </span>
        );
    }

    // Fallback if no AM/PM found (e.g. 24h locale), allow clicking the whole string to switch?
    // Or just render as is for safety.
    return <span>{localString}</span>;
};

export const HistoryTooltip: React.FunctionComponent<IHistoryTooltipProps> = (props) => {
    const { context, value, onChange, fieldName } = props;

    const [history, setHistory] = React.useState<IAuditRecord[]>([]);
    const [loading, setLoading] = React.useState<boolean>(false);
    const [error, setError] = React.useState<string | null>(null);
    const [hasFetched, setHasFetched] = React.useState(false);
    const [isCalloutVisible, setIsCalloutVisible] = React.useState(false);
    const iconRef = React.useRef<HTMLDivElement>(null);

    const fetchHistory = React.useCallback(async (force = false) => {

        let lastKnownValue: string | undefined = undefined;

        if (loading) return;
        if (hasFetched && !force) return;
        setHasFetched(true);
        setLoading(true);
        setError(null);

        const rawEntityId = (context.mode as unknown as { contextInfo: { entityId: string } }).contextInfo?.entityId;
        if (!rawEntityId) {
            setLoading(false);
            return;
        }

        const entityId = rawEntityId.replace(/\{|\}/g, "");

        try {
            const result = await context.webAPI.retrieveMultipleRecords(
                "audit",
                `?$filter=_objectid_value eq ${entityId} &$orderby=createdon desc &$top=30 &$select=createdon,operation,action,changedata,auditid &$expand=userid($select=fullname)`
            );

            const filteredRecords: IAuditRecord[] = [];

            for (const e of result.entities) {
                if (filteredRecords.length >= 20) break;

                let oldValue: string | undefined;
                let newValue: string | undefined;
                let fieldChanged = false;

                const isCreate = e["operation"] === 1;
                const opLabel = isCreate ? "Created" : "Updated";
                const rawChangeData = e["changedata"];

                if (isCreate && value !== null && value !== undefined) {
                    lastKnownValue = value;
                }

                if (rawChangeData) {
                    try {
                        if (rawChangeData.trim().startsWith("{")) {
                            const json = JSON.parse(rawChangeData);
                            if (json.changedAttributes && Array.isArray(json.changedAttributes)) {
                                for (const attr of json.changedAttributes) {
                                    if (attr.logicalName?.toLowerCase() === fieldName.toLowerCase()) {
                                        fieldChanged = true;
                                        oldValue = attr.oldValue === null ? "" : String(attr.oldValue);
                                        newValue = attr.newValue === null ? "(empty)" : String(attr.newValue);
                                        break;
                                    }
                                }
                            }
                        } else {
                            const parser = new DOMParser();
                            const xmlDoc = parser.parseFromString(rawChangeData, "text/xml");
                            const attributes = Array.from(xmlDoc.getElementsByTagName("attribute"));
                            for (const attr of attributes) {
                                if (attr.getAttribute("name")?.toLowerCase() === fieldName.toLowerCase()) {
                                    fieldChanged = true;
                                    const oldNode = attr.getElementsByTagName("oldValue")[0];
                                    const newNode = attr.getElementsByTagName("newValue")[0];
                                    oldValue = oldNode?.textContent || "(empty)";
                                    newValue = newNode?.textContent || "(empty)";
                                    break;
                                }
                            }
                        }
                    } catch (err) {
                        console.error("Error parsing changedata", err);
                    }
                }

                if (fieldChanged) {

                    if (oldValue === undefined && lastKnownValue !== undefined) {
                        oldValue = lastKnownValue;
                    }

                    let oldDisplay = (!oldValue || oldValue === "(empty)") ? "--" : oldValue;
                    let newDisplay = (!newValue || newValue === "(empty)") ? "--" : newValue;

                    lastKnownValue = newValue;

                    const attrMetadata = context.parameters.value.attributes;
                    if (attrMetadata) {
                        const meta = attrMetadata as unknown as IOptionMetadata;
                        if (meta.Type === "OptionSet" || meta.Type === "TwoOptions") {
                            const options = meta.Options;
                            if (options && Array.isArray(options)) {
                                const oldOpt = options.find(o => String(o.Value) === oldValue);
                                const newOpt = options.find(o => String(o.Value) === newValue);
                                if (oldOpt) oldDisplay = oldOpt.Label;
                                if (newOpt) newDisplay = newOpt.Label;
                            }
                        }
                    }

                    filteredRecords.push({
                        createdon: e["createdon"],
                        userid: e["userid"]?.fullname || "System",
                        operation: opLabel,
                        action: e["action@OData.Community.Display.V1.FormattedValue"] || "",
                        auditId: e["auditid"],
                        oldValue: oldDisplay,
                        newValue: newDisplay
                    });
                }
            }

            setHistory(filteredRecords);
        } catch (err: unknown) {
            if (err instanceof Error) {
                setError(err.message);
            } else {
                setError("Error fetching history");
            }
        } finally {
            setLoading(false);
        }
    }, [context, fieldName, value, hasFetched]);

    const onRenderContent = () => {
        if (loading) {
            return (
                <Stack horizontalAlign="center" tokens={{ padding: 20 }}>
                    <Spinner size={SpinnerSize.medium} label="Fetching history..." />
                </Stack>
            );
        }

        if (error) {
            return (
                <Stack tokens={{ padding: 10 }}>
                    <Text variant="small" style={{ color: "#d13438" }}>{error}</Text>
                </Stack>
            );
        }

        if (history.length === 0) {
            return (
                <Stack tokens={{ padding: 15 }} horizontalAlign="center">
                    <Text variant="medium" style={{ fontWeight: 600 }}>No history found</Text>
                    <Text variant="smallPlus" style={{ color: "#605e5c", textAlign: "center" }}>
                        Auditing for "{fieldName}" might be disabled or no changes were recorded yet.
                    </Text>
                </Stack>
            );
        }

        return (
            <Stack tokens={stackTokens} styles={tooltipStackStyles}>
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                    <Text variant="large" style={{ fontWeight: 600, color: "#323130" }}>Field History</Text>
                    <IconButton
                        iconProps={{ iconName: "Cancel" }}
                        onClick={() => setIsCalloutVisible(false)}
                        styles={{ root: { color: "#323130", height: 24, width: 24 } }}
                    />
                </Stack>
                <Separator />
                <Stack tokens={{ childrenGap: 15 }} styles={{ root: { maxHeight: "450px", overflowY: "auto", overflowX: "hidden", paddingRight: "4px" } }}>
                    {history.map((h, index) => (
                        <React.Fragment key={h.auditId}>
                            <ActivityItem
                                activityDescription={[
                                    <div key={1} style={{ color: "#323130", fontSize: "13px" }}>
                                        <span style={{ fontWeight: 600 }}>{h.userid}</span>
                                        <span style={{ color: "#605e5c", margin: "0 4px" }}>&bull;</span>
                                        <span
                                            style={{
                                                fontWeight: 600,
                                                color: h.operation === "Created" ? "#107c10" : "#0078d4"
                                            }}
                                        >
                                            {h.operation}
                                        </span>
                                    </div>,
                                    <div key={2} style={{ color: "#605e5c", fontSize: "11px", marginTop: "1px" }}>
                                        <TimestampToggler dateStr={h.createdon} />
                                    </div>
                                ]}
                                activityIcon={
                                    <div
                                        style={{
                                            display: "flex",
                                            alignItems: "center",
                                            justifyContent: "center",
                                            fontWeight: 600,
                                            fontSize: "10px",
                                            width: "20px",
                                            height: "20px",
                                            borderRadius: "50%",
                                            background: h.operation === "Created" ? "#107c10" : "#0078d4",
                                            color: "#ffffff",
                                            boxShadow: "0 2px 4px rgba(0,0,0,0.1)"
                                        }}
                                    >
                                        {h.operation === "Created" ? "C" : "U"}
                                    </div>
                                }

                                comments={
                                    <Stack tokens={{ childrenGap: 2 }} style={{ marginTop: 2 }}>
                                        <Stack horizontal tokens={{ childrenGap: 5 }} verticalAlign="center">
                                            <Text variant="small" style={{ fontWeight: 600, color: "#a4262c", whiteSpace: "nowrap" }}>Old Value:</Text>
                                            <Text variant="small" style={{ fontStyle: "italic", userSelect: "text", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", maxWidth: "160px", display: "block" }} title={h.oldValue || ""}>{h.oldValue}</Text>
                                        </Stack>
                                        <Stack horizontal tokens={{ childrenGap: 5 }} verticalAlign="center">
                                            <Text variant="small" style={{ fontWeight: 600, color: "#107c10", whiteSpace: "nowrap" }}>New Value:</Text>
                                            <Text variant="small" style={{ fontStyle: "italic", userSelect: "text", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", maxWidth: "160px", display: "block" }} title={h.newValue || ""}>{h.newValue}</Text>
                                        </Stack>
                                    </Stack>
                                }
                            />
                            {index < history.length - 1 && (
                                <Separator styles={{ root: { padding: 0 } }} />
                            )}
                        </React.Fragment>
                    ))}
                </Stack>
                <Separator />
                <Text variant="tiny" style={{ color: "#a19f9d", textAlign: "right" }}>
                    Showing last {history.length} changes
                </Text>
            </Stack>
        );
    };

    const [localValue, setLocalValue] = React.useState(value || "");

    React.useEffect(() => {
        setLocalValue(value || "");
    }, [value]);

    return (
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} style={{ width: "100%" }}>
            <TextField
                value={localValue}
                onChange={(e, v) => setLocalValue(v || "")}
                onBlur={() => onChange(localValue)}
                onKeyDown={(e) => e.stopPropagation()}
                borderless
                autoComplete="off"
                styles={{
                    root: { flexGrow: 1 },
                    fieldGroup: {
                        background: "transparent",
                        borderBottom: "1px solid #deecf9",
                        selectors: {
                            ":after": { borderBottomColor: "#0078d4" }
                        }
                    },
                    field: { fontSize: "14px", color: "#323130" }
                }}
            />
            <div
                ref={iconRef}
                onClick={() => {
                    const newState = !isCalloutVisible;
                    setIsCalloutVisible(newState);
                    if (newState) fetchHistory(true);
                }}
                style={{
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    width: "28px",
                    height: "28px",
                    borderRadius: "50%",
                    cursor: "pointer",
                    transition: "background 0.2s",
                    background: isCalloutVisible ? "#deecf9" : "#f3f2f1",
                    color: "#0078d4"
                }}
                onMouseOver={(e) => !isCalloutVisible && (e.currentTarget.style.background = "#deecf9")}
                onMouseOut={(e) => !isCalloutVisible && (e.currentTarget.style.background = "#f3f2f1")}
            >
                <Icon iconName="History" styles={{ root: { fontSize: "14px" } }} />
            </div>
            {isCalloutVisible && (
                <Callout
                    target={iconRef}
                    onDismiss={() => setIsCalloutVisible(false)}
                    directionalHint={DirectionalHint.bottomRightEdge}
                    gapSpace={10}
                    beakWidth={10}
                    styles={{
                        beak: { background: "#fff" },
                        beakCurtain: { background: "transparent" },
                        calloutMain: { borderRadius: "8px", border: "none", outline: "none" }
                    }}
                    setInitialFocus
                >
                    {onRenderContent()}
                </Callout>
            )}
        </Stack>
    );
};
