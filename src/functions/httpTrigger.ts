import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import * as AppInsights from 'applicationinsights';
import { MyItem } from "../common/models";

export async function httpTrigger(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
  // Initialize Application Insights, setting AutoDependencyCorrelation to true so that all logs for each run are correlated together in App Insights. 
  AppInsights.setup(process.env.APPLICATIONINSIGHTS_CONNECTION_STRING).setAutoDependencyCorrelation(true);
  AppInsights.start();

  // Set the default response to a 200 with an empty body.
  let retVal: HttpResponseInit = { status: 200, body: "" };

  // If the request is a GET or DELETE, then we need to get the ID from the query string.
  if (request.method == "GET" || request.method == "DELETE") {
    const id = request.query.get("id");
    if (id != null) {

    } else {
      retVal = { status: 400, body: "Invalid request" };
    }
  } else if (request.method == "POST" || request.method == "PATCH") {
    // If the request is a POST or PATCH, then we need to get the items payload from the body.
    const item: MyItem = (await request.json()) as MyItem;
    if (item != null) {

    } else {
      retVal = { status: 400, body: "Invalid request" };
    }
  }
  // Return the appropriate response to the requestor.
  return retVal;
};

app.http('httpTrigger', {
  methods: ['GET', 'POST', 'PATCH', 'DELETE'],
  authLevel: 'anonymous',
  handler: httpTrigger
});
