// Export function to make it available for import in other files
export async function setEmailSenderToQueue(ctx: Xrm.Events.EventContext) {
    let formContext = ctx.getFormContext();
    // Only run if form type is create or if sender is empty (null)
    if (formContext.ui.getFormType() === XrmEnum.FormType.Create || formContext.getAttribute("from").getValue() === null) {
        let regarding = formContext.getAttribute("regardingobjectid").getValue() as Xrm.LookupValue[];
        // Validate that email is set regarding and entity type is incident. Last part can be removed to support any parent entity
        if (regarding !== null && regarding[0].entityType.toLowerCase() === "incident") {
            // Retrieve active queue item with queue reference. Alternatively add more logic to verify that the queue is activated for email
            var queue = await Xrm.WebApi.retrieveMultipleRecords(
                "queueitem",
                "?$select=_queueid_value&$filter=statecode eq 0 and _objectid_value eq " + regarding[0].id
            ).then(getEntityReferenceFromMultipleResult);
            
            // Set sender to queue if exists
            if (queue !== null) {
                formContext.getAttribute("from").setValue([queue])
            }
        }
    }
}

async function getEntityReferenceFromMultipleResult(results: Xrm.RetrieveMultipleResult): Promise<Xrm.LookupValue> {
    // Return lookup value of queue if exactly 1 active queueitem is found
    if (results.entities.length === 1) {
        let queue = {} as Xrm.LookupValue;
        queue.entityType = results.entities[0]["_queueid_value@Microsoft.Dynamics.CRM.lookuplogicalname"];
        queue.id = results.entities[0]._queueid_value;
        queue.name = results.entities[0]["_queueid_value@OData.Community.Display.V1.FormattedValue"];

        return queue;
    }
    // Throw error if more than 1 active queueitem is found
    if (results.entities.length > 1)
    {
        throw "Number of active queue items is more than 1";
    }
    // Return null by default
    return null;
}