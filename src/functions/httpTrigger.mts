import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import AppInsights from 'applicationinsights';
import { MyItem } from "../common/models.mjs";
import { pnpjs } from "../common/pnpjsService.mjs";

export async function httpTrigger(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
  const LOG_SOURCE = "httpTrigger";
  // Initialize Application Insights, setting AutoDependencyCorrelation to true so that all logs for each run are correlated together in App Insights. 
  AppInsights.setup(process.env.APPLICATIONINSIGHTS_CONNECTION_STRING).setAutoDependencyCorrelation(true);
  // set the current tagged app & version of the function app
  AppInsights.defaultClient.commonProperties = { 'app.name': "PnPjsAzureFunctionSample" };
  AppInsights.defaultClient.context.tags[AppInsights.defaultClient.context.keys.applicationVersion] = "1.0.0";

  // DEV ONLY: set the operation id to the current invocation id
  //AppInsights.defaultClient.context.tags[AppInsights.defaultClient.context.keys.operationId] = context.invocationId; 

  // start the client
  AppInsights.start();

  AppInsights.defaultClient.trackEvent({
    name: `${LOG_SOURCE}/request`,
    properties: {
      source: LOG_SOURCE,
      requestBody: JSON.stringify(request.body),
      requestQuery: JSON.stringify(request.query)
    }
  });

  // Set the default response to a 200 with an empty body.
  let retVal: HttpResponseInit = { status: 200, body: "" };

  try {
    // If the request is a GET or DELETE, then we need to get the ID from the query string.
    if (request.method == "GET" || request.method == "DELETE") {
      const id = request.query.get("id");
      if (id != null) {
        const ready = await pnpjs.Init();
        if (ready) {
          const result = await pnpjs.GetListItem(id);
          if (result != null) {
            retVal = { status: 200, body: JSON.stringify(result) };
            AppInsights.defaultClient.trackTrace({
              message: 'Found item',
              properties: {
                source: LOG_SOURCE,
                request_type: "GET"
              },
              severity: AppInsights.Contracts.SeverityLevel.Verbose
            });
          } else {
            retVal = { status: 400, body: "Item not found." };
            AppInsights.defaultClient.trackTrace({
              message: 'Item not found',
              properties: {
                source: LOG_SOURCE,
                request_type: "GET"
              },
              severity: AppInsights.Contracts.SeverityLevel.Verbose
            });
          }
        }
      } else {
        retVal = { status: 400, body: "Invalid request" };
        AppInsights.defaultClient.trackTrace({
          message: 'Invalid Request',
          properties: {
            source: LOG_SOURCE,
            request_type: "GET"
          },
          severity: AppInsights.Contracts.SeverityLevel.Verbose
        });
      }
    } else if (request.method == "POST" || request.method == "PATCH") {
      // If the request is a POST or PATCH, then we need to get the items payload from the body.
      const item: MyItem = (await request.json()) as MyItem;
      if (item != null) {
        const ready = await pnpjs.Init();
        if (ready) {
          //Will update item
        }
      } else {
        retVal = { status: 400, body: "Invalid request" };
        AppInsights.defaultClient.trackTrace({
          message: 'Invalid Request',
          properties: {
            source: LOG_SOURCE,
            request_type: "POST"
          },
          severity: AppInsights.Contracts.SeverityLevel.Verbose
        });
      }
    }
  } catch (err) {
    AppInsights.defaultClient.trackException({
      exception: err,
      severity: AppInsights.Contracts.SeverityLevel.Critical,
      properties: { source: LOG_SOURCE, method: "httpTrigger" }
    });
  }

  // Return the appropriate response to the requestor.
  return retVal;
};

app.http('httpTrigger', {
  methods: ['GET', 'POST', 'PATCH', 'DELETE'],
  authLevel: 'anonymous',
  handler: httpTrigger
});
