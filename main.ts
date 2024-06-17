type Criteria = {
    field: CriteriaField;
    value: string;
};

enum Action {
    LABEL = 'label',
    DELETE = 'delete',
    MARK_READ = 'markRead',

}

type Settings = {
    enableAutoApply: boolean;
    autoApplyIntervalHours: number;
    mostRecentMails: number;

};

enum CriteriaField {
    FROM = "from",
    SUBJECT = "subjectContains",
    BODY = "bodyContains",
    ATTACHEMENT = "hasAttachment",
    ATTACHEMENT_NAME = "attachmentName",
    TO = "to",
    CC = "cc",
    BCC = "bcc",
}

type Rule = {
    id: number;
    name: string;
    description: string;
    criteria: Criteria;
    action: Action;
    value: string;
    enabled: boolean;
    priority: number;
};

/**
 {
 "id": "000",
 "criteria":
 {
 "field": "from",
 "value": "newsletter@example.com"
 },
 "action": "label",
 "value": "Promotions",
 "enabled": true,
 "priority": 1
 }
 */

function addNewRule(
    criteria: Criteria, action: Action, value: string, name: string, description: string, enabled: boolean, priority: number
): void {
    let scriptProperties = PropertiesService.getScriptProperties();
    let rules: Array<Rule> = JSON.parse(scriptProperties.getProperty('rules') || '[]');

    let id: number = rules.reduce((max, rule) => Math.max(max, rule.id), 0) + 1;

    let ruleNew: Rule = {
        id: id,
        name: name || "Unnamed Rule", // Default name if not provided
        description: description || "", // Keep as empty string if not provided
        criteria: criteria || {field: CriteriaField.FROM, value: ""}, // Default criteria if not provided
        action: action || Action.LABEL, // Default action if not provided
        value: value || "", // Keep as empty string if not provided
        enabled: enabled !== undefined ? enabled : true, // Default to true if not provided
        priority: priority || 0 // Default to 0 if not provided
    };

    rules.push(ruleNew);
    scriptProperties.setProperty('rules', JSON.stringify(rules));
}


function getRules(): Rule[] {
    let scriptProperties = PropertiesService.getScriptProperties();
    try {
        let rules: Rule[] = JSON.parse(scriptProperties.getProperty('rules') || '[]');
        return rules;
    } catch (error) {
        console.error("Failed to parse rules:", error);
        return [];  // Return an empty array on error
    }
}

function confirmRuleDeletion(e: any): GoogleAppsScript.Card_Service.Card {
    const ruleId = e.parameters.ruleId;

    let cardBuilder = CardService.newCardBuilder();
    cardBuilder.setHeader(CardService.newCardHeader().setTitle("Confirm Deletion"));

    let section = CardService.newCardSection();
    section.addWidget(CardService.newTextParagraph().setText("Are you sure you want to delete this rule? This action cannot be undone."));

    section.addWidget(CardService.newTextButton()
        .setText("Confirm Delete")
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setBackgroundColor("#FF0000")
        .setOnClickAction(CardService.newAction().setFunctionName("deleteRule").setParameters({ruleId: ruleId})));

    section.addWidget(CardService.newTextButton()
        .setText("Cancel")
        .setOnClickAction(CardService.newAction().setFunctionName("showEditRuleForm").setParameters({ruleId: ruleId})));

    cardBuilder.addSection(section);
    return cardBuilder.build();
}

function deleteRule(e: any): GoogleAppsScript.Card_Service.Card {
    const ruleId = parseInt(e.parameters.ruleId, 10);
    let rules = getRules();
    rules = rules.filter(rule => rule.id !== ruleId);

    let scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('rules', JSON.stringify(rules));

    let cardBuilder = CardService.newCardBuilder();
    cardBuilder.setHeader(CardService.newCardHeader().setTitle("Rule Deleted"));
    cardBuilder.addSection(CardService.newCardSection().addWidget(CardService.newTextParagraph().setText("The rule has been successfully deleted.")));

    cardBuilder.addSection(
        CardService.newCardSection().addWidget(
            CardService.newTextButton()
                .setText("Back to Rules")
                .setOnClickAction(CardService.newAction().setFunctionName("loadAddOn"))
        )
    );

    return cardBuilder.build();
}

function getEnums() {
    return {
        Action: Action,
        CriteriaField: CriteriaField
    };
}


function debug(): void {
    console.log("Debugging Actions and CriteriaFields:");
    console.log(getEnums());
    console.log("Current Rules:");
    let rules = getRules();
    if (rules.length === 0) {
        console.log("No rules defined.");
    } else {
        console.log(rules);
    }
}


function getFormattedRulesWidgets(): GoogleAppsScript.Card_Service.Widget[] {
    let rules = getRules();
    let widgets: GoogleAppsScript.Card_Service.Widget[] = [];

    if (rules.length === 0) {
        widgets.push(CardService.newTextParagraph().setText("No rules have been defined."));
        return widgets;
    }

    rules.forEach(rule => {
        const textColor = rule.enabled ? "#00FF00" : "#FF0000"; // Green for enabled, red for disabled
        const widget = CardService.newTextButton()
            .setText(`${rule.name} (Enabled: ${rule.enabled ? 'Yes' : 'No'})`)
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED) // Ensure this is valid in your context
            .setBackgroundColor(textColor) // SetBackgroundColor might not be valid; usually not supported in CardService
            .setOnClickAction(CardService.newAction().setFunctionName("showEditRuleForm").setParameters({ruleId: rule.id.toString()}));
        widgets.push(widget);
    });

    return widgets;
}

function buildAddRuleForm(action: string | null): GoogleAppsScript.Card_Service.Card {
    let cardBuilder = CardService.newCardBuilder();
    cardBuilder.setHeader(CardService.newCardHeader().setTitle("Create New Rule"));

    let section = CardService.newCardSection();

    // Text input for Rule Name
    section.addWidget(CardService.newTextInput()
        .setFieldName("ruleName")
        .setTitle("Rule Name"));

    // Text input for Rule Description
    section.addWidget(CardService.newTextInput()
        .setMultiline(true)
        .setFieldName("ruleDescription")
        .setTitle("Rule Description"));

    // Dropdown for Criteria Field
    let criteriaFieldDropdown = CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.DROPDOWN)
        .setTitle("Criteria Field")
        .setFieldName("criteriaField");
    for (const [key, value] of Object.entries(CriteriaField)) {
        criteriaFieldDropdown.addItem(key, value, value === CriteriaField.FROM);
    }
    section.addWidget(criteriaFieldDropdown);


    // Text input for Criteria Value
    section.addWidget(CardService.newTextInput()
        .setFieldName("criteriaValue")
        .setTitle("Criteria Value"));

    // Dropdown for Actions
    let actionDropdown = CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.DROPDOWN)
        .setTitle("Action")
        .setFieldName("action")
        .setOnChangeAction(CardService.newAction().setFunctionName("handleActionChange"));
    for (const [key, value] of Object.entries(Action)) {
        actionDropdown.addItem(key, value, value === action);
    }

    section.addWidget(actionDropdown);

    // Conditional input for Action Value based on selected action
    if (action === "label") {
        const labels = GmailApp.getUserLabels();
        const labelDropdown = CardService.newSelectionInput()
            .setType(CardService.SelectionInputType.DROPDOWN)
            .setTitle("Label")
            .setFieldName("actionValue");
        labels.forEach(label => {
            labelDropdown.addItem(label.getName(), label.getName(), false);
        });
        section.addWidget(labelDropdown);
    } else {
        section.addWidget(CardService.newTextInput()
            .setFieldName("actionValue")
            .setTitle("Action Value"));
    }

    // Checkbox for Enabled
    section.addWidget(CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.CHECK_BOX)
        .setTitle("Enabled")
        .setFieldName("enabled")
        .addItem("Yes", "true", true));

    // Text input for Priority
    section.addWidget(CardService.newTextInput()
        .setValue("1")
        .setFieldName("priority")
        .setTitle("Priority"));

    // Button to submit the form
    section.addWidget(CardService.newTextButton()
        .setText("Create Rule")
        .setOnClickAction(CardService.newAction()
            .setFunctionName("handleCreateRule")));

    cardBuilder.addSection(section);

    return cardBuilder.build();
}

function showAddRuleForm() {

    return buildAddRuleForm(null);
}

function handleActionChange(e: any): GoogleAppsScript.Card_Service.Card {
    const action = e.formInput.action;
    return buildAddRuleForm(action);
}

function handleCreateRule(e: any): GoogleAppsScript.Card_Service.Card {
    const formInputs = e.formInput;

    addNewRule({
            field: formInputs.criteriaField as CriteriaField,
            value: formInputs.criteriaValue
        },
        formInputs.action as Action,
        formInputs.actionValue,
        formInputs.ruleName,
        formInputs.ruleDescription,
        formInputs.enabled === 'true',
        parseInt(formInputs.priority));

    return CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader().setTitle("Rule Created Successfully"))
        .addSection(
            CardService
                .newCardSection()
                .addWidget(
                    CardService
                        .newTextButton()
                        .setText("Back")
                        .setOnClickAction(
                            CardService
                                .newAction()
                                .setFunctionName("loadAddOn")
                        )))
        .build();
}

function handleSaveRuleChanges(e: any): GoogleAppsScript.Card_Service.Card {
    const ruleId = parseInt(e.parameters.ruleId, 10);
    const formInput = e.formInput;

    let rules = getRules();
    let ruleIndex = rules.findIndex(r => r.id === ruleId);
    if (ruleIndex !== -1) {
        // Update the rule with new values from the form
        rules[ruleIndex] = {
            ...rules[ruleIndex],
            name: formInput.ruleName,
            description: formInput.ruleDescription,
            criteria: {
                field: formInput.criteriaField as CriteriaField,
                value: formInput.criteriaValue
            },
            action: formInput.action as Action,
            value: formInput.actionValue,
            enabled: formInput.enabled === 'true',
            priority: parseInt(formInput.priority, 10)
        };

        // Save updated rules back to storage
        let scriptProperties = PropertiesService.getScriptProperties();
        scriptProperties.setProperty('rules', JSON.stringify(rules));

        return CardService.newCardBuilder()
            .setHeader(CardService.newCardHeader().setTitle("Success"))
            .addSection(CardService.newCardSection().addWidget(CardService.newTextParagraph().setText("Rule updated successfully")))
            .addSection(
                CardService
                    .newCardSection()
                    .addWidget(
                        CardService
                            .newTextButton()
                            .setText("Back")
                            .setOnClickAction(
                                CardService
                                    .newAction()
                                    .setFunctionName("loadAddOn")
                            )))
            .build();
    } else {
        return CardService.newCardBuilder()
            .setHeader(CardService.newCardHeader().setTitle("Error"))
            .addSection(CardService.newCardSection().addWidget(CardService.newTextParagraph().setText("Failed to update the rule")))
            .addSection(
                CardService
                    .newCardSection()
                    .addWidget(
                        CardService
                            .newTextButton()
                            .setText("Back")
                            .setOnClickAction(
                                CardService
                                    .newAction()
                                    .setFunctionName("loadAddOn")
                            )))
            .build();
    }
}


function showEditRuleForm(e: any): GoogleAppsScript.Card_Service.Card {
    const ruleId = parseInt(e.parameters.ruleId, 10);
    const rules = getRules();
    const rule = rules.find(r => r.id === ruleId);

    if (!rule) {
        return CardService.newCardBuilder()
            .setHeader(CardService.newCardHeader().setTitle("Error"))
            .addSection(CardService.newCardSection().addWidget(CardService.newTextParagraph().setText("Rule not found")))
            .build();
    }

    let cardBuilder = CardService.newCardBuilder();
    cardBuilder.setHeader(CardService.newCardHeader().setTitle("Edit Rule"));

    let section = CardService.newCardSection();

    // Rule Name
    section.addWidget(CardService.newTextInput()
        .setFieldName("ruleName")
        .setTitle("Rule Name")
        .setValue(rule.name));

    // Rule Description
    section.addWidget(CardService.newTextInput()
        .setMultiline(true)
        .setFieldName("ruleDescription")
        .setTitle("Description")
        .setValue(rule.description || '')); // Handle potentially undefined values

    // Criteria Field
    const criteriaFieldDropdown = CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.DROPDOWN)
        .setTitle("Criteria")
        .setFieldName("criteriaField");
    for (const [key, value] of Object.entries(CriteriaField)) {
        criteriaFieldDropdown.addItem(key, value, value === rule.criteria.field);
    }
    section.addWidget(criteriaFieldDropdown);

    // Criteria Value
    section.addWidget(CardService.newTextInput()
        .setFieldName("criteriaValue")
        .setTitle("Criteria Value")
        .setValue(rule.criteria.value));

    // Action Dropdown
    const actionDropdown = CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.DROPDOWN)
        .setTitle("Action")
        .setFieldName("action");
    for (const [key, value] of Object.entries(Action)) {
        actionDropdown.addItem(key, value, value === rule.action);
    }
    section.addWidget(actionDropdown);

    // Rule Action Value
    section.addWidget(CardService.newTextInput()
        .setFieldName("actionValue")
        .setTitle("Action Value")
        .setValue(rule.value));

    // Enabled
    section.addWidget(CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.CHECK_BOX)
        .setTitle("Enabled")
        .setFieldName("enabled")
        .addItem("Yes", "true", rule.enabled));

    // Priority
    section.addWidget(CardService.newTextInput()
        .setFieldName("priority")
        .setTitle("Priority")
        .setValue(rule.priority.toString()));

    // Button to save the changes
    section.addWidget(CardService.newTextButton()
        .setText("Save Changes")
        .setBackgroundColor("#00FF00")
        .setOnClickAction(CardService.newAction()
            .setFunctionName("handleSaveRuleChanges")
            .setParameters({ruleId: rule.id.toString()})))
        .addWidget(CardService.newTextButton()
            .setText("Delete Rule")
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setBackgroundColor("#FF0000") // Optional: Red color to highlight the deletion action
            .setOnClickAction(CardService.newAction().setFunctionName("confirmRuleDeletion").setParameters({ruleId: rule.id.toString()})));

    cardBuilder.addSection(section);
    cardBuilder.addSection(
        CardService
            .newCardSection()
            .addWidget(
                CardService
                    .newTextButton()
                    .setText("Back")
                    .setOnClickAction(
                        CardService
                            .newAction()
                            .setFunctionName("loadAddOn")
                    )));

    return cardBuilder.build();
}


function loadAddOn() {
    let card = CardService.newCardBuilder();

    card.setName("Gmail Rules")
        .setHeader(CardService.newCardHeader().setTitle("Gmail Rules"));

    // Create one section for all rules
    let section = CardService.newCardSection()
        .setHeader("Rules"); // Optional header for the section

    // Retrieve and add rule widgets to the section
    let ruleWidgets = getFormattedRulesWidgets();
    ruleWidgets.forEach(widget => section.addWidget(widget));

    // Add button to create new rule in the same section
    section.addWidget(
        CardService.newTextButton()
            .setText("Add Rule")
            .setOnClickAction(CardService.newAction().setFunctionName("showAddRuleForm"))
    );


    card.addSection(
        CardService.newCardSection().addWidget(
            CardService.newTextButton()
                .setText("Settings")
                .setOnClickAction(CardService.newAction().setFunctionName("showSettingsForm"))
        ).addWidget(
            CardService.newTextButton()
                .setText("Manually Trigger")
                .setOnClickAction(CardService.newAction().setFunctionName("applyRules"))
        )
    );
    // Add the populated section to the card

    card.addSection(section);

    return card.build();
}


// function applyRules(): void {
//     const settings = getSettings();
//     const rules = getRules();
//
//     console.log('Applying rules...');
//     console.log('Settings:', JSON.stringify(settings));
//     console.log('Number of rules:', rules.length);
//
//     if (rules.length === 0) {
//         console.log('No rules to apply.');
//         return;
//     }
//
//     const threads = GmailApp.getInboxThreads(0, settings.mostRecentMails);
//     console.log('Number of threads to process:', threads.length);
//
//     threads.forEach(thread => {
//         console.log('Processing thread with ID:', thread.getId());
//         const messages = thread.getMessages();
//         console.log('Number of messages in thread:', messages.length);
//
//         messages.forEach(message => {
//             console.log('Processing message with ID:', message.getId());
//             rules.forEach(rule => {
//                 console.log('Evaluating rule:', JSON.stringify(rule));
//                 if (rule.enabled && matchesRule(message, rule)) {
//                     console.log('Rule matched. Executing action...');
//                     executeRuleAction(message, rule);
//                 } else {
//                     console.log('Rule did not match or is not enabled.');
//                 }
//             });
//         });
//     });
//
//     console.log('Finished applying rules.');
// }
//
//
// function matchesRule(message: GoogleAppsScript.Gmail.GmailMessage, rule: Rule): boolean {
//     console.log(`Checking if message with ID ${message.getId()} matches rule ${rule.id}`);
//     switch (rule.criteria.field) {
//         case CriteriaField.FROM:
//             const fromMatch = message.getFrom().includes(rule.criteria.value);
//             console.log(`Criteria FROM matched: ${fromMatch} ${message.getFrom()}`);
//             return fromMatch;
//         case CriteriaField.TO:
//             const toMatch = message.getTo().includes(rule.criteria.value);
//             console.log(`Criteria TO matched: ${toMatch} ${message.getTo()}`);
//             return toMatch;
//         case CriteriaField.CC:
//             const ccMatch = message.getCc().includes(rule.criteria.value);
//             console.log(`Criteria CC matched: ${ccMatch} ${message.getCc()}`);
//             return ccMatch;
//         case CriteriaField.BCC:
//             const bccMatch = message.getBcc().includes(rule.criteria.value);
//             console.log(`Criteria BCC matched: ${bccMatch} ${message.getBcc()}`);
//             return bccMatch;
//         default:
//             console.log('Criteria field not recognized:', rule.criteria.field);
//             return false;
//     }
// }
//
// function executeRuleAction(message: GoogleAppsScript.Gmail.GmailMessage, rule: Rule): void {
//     console.log(`Executing action for message with ID ${message.getId()} based on rule ${rule.id}`);
//     switch (rule.action) {
//         case Action.LABEL:
//             const labelName = rule.value;
//             let label = GmailApp.getUserLabelByName(labelName);
//             if (!label) {
//                 try {
//                     label = GmailApp.createLabel(labelName);
//                     console.log(`Label '${labelName}' created.`);
//                 } catch (error) {
//                     console.error(`Failed to create label '${labelName}':`, error);
//                     return;
//                 }
//             } else {
//                 console.log(`Label '${labelName}' found.`);
//             }
//
//
//             label.addToThread(message.getThread());
//             message.getThread().moveToArchive();
//             console.log(`Label '${rule.value}' added to thread with ID ${message.getThread().getId()}`);
//             break;
//         case Action.DELETE:
//             message.moveToTrash();
//             console.log(`Message with ID ${message.getId()} moved to trash`);
//             break;
//         case Action.MARK_READ:
//             message.markRead();
//             console.log(`Message with ID ${message.getId()} marked as read`);
//             break;
//         default:
//             console.log('Action not recognized:', rule.action);
//             break;
//     }
// }

function applyRules(): void {
    const settings = getSettings();
    const rules = getRules();

    console.log('Applying rules...');
    if (rules.length === 0) {
        console.log('No rules to apply.');
        return;
    }

    const threads = GmailApp.getInboxThreads(0, settings.mostRecentMails);
    if (threads.length === 0) {
        console.log('No threads to process.');
        return;
    }

    // Log only the count of threads to reduce log clutter
    console.log('Number of threads to process:', threads.length);

    // Create a mapping of rules for quick access
    const rulesMap = new Map<number, Rule>(rules.map(rule => [rule.id, rule]));

    // Process all threads in batches
    threads.forEach(thread => {
        const messages = thread.getMessages();
        messages.forEach(message => {
            rules.forEach(rule => {
                if (rule.enabled && matchesRule(message, rule)) {
                    executeRuleAction(thread, message, rule);
                }
            });
        });
    });

    console.log('Finished applying rules.');
}

function matchesRule(message: GoogleAppsScript.Gmail.GmailMessage, rule: Rule): boolean {
    // Simplified logging for clarity
    switch (rule.criteria.field) {
        case CriteriaField.FROM:
            return message.getFrom().includes(rule.criteria.value);
        case CriteriaField.TO:
            return message.getTo().includes(rule.criteria.value);
        case CriteriaField.CC:
            return message.getCc().includes(rule.criteria.value);
        case CriteriaField.BCC:
            return message.getBcc().includes(rule.criteria.value);
        case CriteriaField.SUBJECT:
            return message.getSubject().includes(rule.criteria.value);
        case CriteriaField.BODY:
            return message.getBody().includes(rule.criteria.value);
        case CriteriaField.ATTACHEMENT:
            // Placeholder: Implement specific logic to check for attachments
            return message.getAttachments().length > 0;
        case CriteriaField.ATTACHEMENT_NAME:
            // Placeholder: Check attachment names
            return message.getAttachments().some(attachment => attachment.getName().includes(rule.criteria.value));
        default:
            console.error('Unrecognized criteria field:', rule.criteria.field);
            return false;
    }
}

function executeRuleAction(thread: GoogleAppsScript.Gmail.GmailThread, message: GoogleAppsScript.Gmail.GmailMessage, rule: Rule): void {
    console.log(`Executing rule: ${rule.name} for message ID: ${message.getId()}`);
    switch (rule.action) {
        case Action.LABEL:
            let label = GmailApp.getUserLabelByName(rule.value);
            if (!label) label = GmailApp.createLabel(rule.value);
            label.addToThread(thread);
            thread.moveToArchive();
            break;
        case Action.DELETE:
            message.moveToTrash();
            break;
        case Action.MARK_READ:
            message.markRead();
            break;
        default:
            console.error('Action not recognized:', rule.action);
            break;
    }
}

function addTrigger(): void {
    const settings = getSettings();

    // Validate settings
    if (settings.autoApplyIntervalHours <= 0) {
        console.error('Invalid auto apply interval hours:', settings.autoApplyIntervalHours);
        return;
    }

    // First, delete existing triggers to avoid duplicates
    ScriptApp.getProjectTriggers().forEach(trigger => {
        if (trigger.getHandlerFunction() === 'applyRules') {
            ScriptApp.deleteTrigger(trigger);
        }
    });

    // Check if auto-apply is enabled before setting up a new trigger
    if (settings.enableAutoApply) {
        ScriptApp.newTrigger('applyRules')
            .timeBased()
            .everyHours(settings.autoApplyIntervalHours)
            .create();
        console.log(`Trigger created to run every ${settings.autoApplyIntervalHours} hours.`);
    } else {
        console.log('Auto-apply is disabled. No trigger was created.');
    }
}

// Function to get the current settings
function getSettings(): Settings {
    const scriptProperties = PropertiesService.getScriptProperties();
    const settingsJson = scriptProperties.getProperty('settings');
    return settingsJson ? JSON.parse(settingsJson) as Settings : {
        enableAutoApply: false,
        autoApplyIntervalHours: 1,  // Default to every 1 hour
        mostRecentMails: 50  // Default to the most recent 50 mails
    };
}

// Function to save settings
function setSettings(settings: Settings): void {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('settings', JSON.stringify(settings));
}

function removeTrigger(): void {
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
        if (trigger.getHandlerFunction() === 'applyRules') {
            ScriptApp.deleteTrigger(trigger);
        }
    }
}

function showSettingsForm(): GoogleAppsScript.Card_Service.Card {
    let settings = getSettings();
    let cardBuilder = CardService.newCardBuilder();

    cardBuilder.setHeader(CardService.newCardHeader().setTitle("Settings"));

    let section = CardService.newCardSection();
    section.addWidget(
        CardService.newSelectionInput()
            .setType(CardService.SelectionInputType.CHECK_BOX)
            .setTitle("Enable Auto Apply")
            .setFieldName("enableAutoApply")
            .addItem("Yes", "true", settings.enableAutoApply)
    )
        .addWidget(
            CardService.newTextInput()
                .setFieldName("autoApplyIntervalHours")
                .setTitle("Auto Apply Interval (hours)")
                .setValue(settings.autoApplyIntervalHours.toString())
        )
        .addWidget(
            CardService.newTextInput()
                .setFieldName("mostRecentMails")
                .setTitle("Most Recent Mails")
                .setValue(settings.mostRecentMails.toString())
        )
        .addWidget(
            CardService.newTextButton()
                .setText("Save Settings")
                .setOnClickAction(CardService.newAction().setFunctionName("handleSaveSettings"))
        )
        .addWidget(CardService.newKeyValue()
            .setTopLabel("Trigger Status")
            .setContent(isTriggerCreated() ? "Trigger is active" : "Trigger is not active")
            .setIcon(isTriggerCreated() ? CardService.Icon.CLOCK : CardService.Icon.STAR));

    return cardBuilder.addSection(section).build();
}

function handleSaveSettings(e: any): GoogleAppsScript.Card_Service.Card {
    const formInputs = e.formInputs;

    // Safely parse integers with fallbacks
    const autoApplyIntervalHours = parseInt(formInputs.autoApplyIntervalHours) || 60; // Default to 60 if parsing fails
    const mostRecentMails = parseInt(formInputs.mostRecentMails) || 50; // Default to 50 if parsing fails

    // Ensure that provided settings are logical
    if (![1, 2, 4, 6, 8, 12].includes(autoApplyIntervalHours) || mostRecentMails <= 0) {
        return createErrorCard("Please enter 1, 2, 4, 6, 8, 12 for interval and positive number for mail count.");
    }

    // Retrieve existing settings
    const currentSettings = getSettings();
    if (formInputs.enableAutoApply == undefined) {
        formInputs.enableAutoApply = [];
    }
    let settings: Settings = {
        enableAutoApply: formInputs.enableAutoApply.includes('true') ?? false,
        autoApplyIntervalHours: autoApplyIntervalHours,
        mostRecentMails: mostRecentMails
    };

    // Save the updated settings
    setSettings(settings);

    addTrigger();

    // Return a card indicating success and providing a navigation option
    let returnCard: GoogleAppsScript.Card_Service.Card = createSuccessCard().build();
    return returnCard;
}

function createSuccessCard(): GoogleAppsScript.Card_Service.CardBuilder {
    return CardService.newCardBuilder()
        .addSection(
            CardService.newCardSection().addWidget(
                CardService.newTextParagraph().setText("Settings saved successfully.")
            )
        ).addSection(
            CardService.newCardSection().addWidget(
                CardService.newTextButton()
                    .setText("Back to Settings")
                    .setOnClickAction(
                        CardService.newAction().setFunctionName("showSettingsForm")
                    )
            )
        );
}

function createErrorCard(errorMessage: string): GoogleAppsScript.Card_Service.Card {
    return CardService.newCardBuilder()
        .addSection(
            CardService.newCardSection().addWidget(
                CardService.newTextParagraph().setText(errorMessage)
            )
        ).addSection(
            CardService.newCardSection().addWidget(
                CardService.newTextButton()
                    .setText("Try Again")
                    .setOnClickAction(
                        CardService.newAction().setFunctionName("showSettingsForm")
                    )
            )
        ).build();
}

function isTriggerCreated(): boolean {
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
        if (trigger.getHandlerFunction() === 'applyRules') {
            return true;
        }
    }
    return false;
}

function handleRemoveTrigger(e: any): GoogleAppsScript.Card_Service.Card {
    removeTrigger();

    return CardService.newCardBuilder()
        .addSection(
            CardService.newCardSection().addWidget(
                CardService.newTextParagraph().setText("Trigger removed successfully.")
            )
        ).addSection(
            CardService.newCardSection().addWidget(
                CardService.newTextButton()
                    .setText("Back to Settings")
                    .setOnClickAction(
                        CardService.newAction().setFunctionName("showSettingsPage")
                    )
            )
        ).build();
}