#r "Newtonsoft.Json"
using System;
using System.Net;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;

public static string HtmlToText(string HTMLCode)
{
    // Feature descriptions often have lots of rich text given in HTML, making the card very long.
    // If the card is too long, the card will fail silently and not show up in Teams.
    // This function tries to strip away the really egregious HTML while maintaining readability.

    // Remove new lines since they are not visible in HTML
    HTMLCode = HTMLCode.Replace("\n", " ");
    
    // Remove tab spaces
    HTMLCode = HTMLCode.Replace("\t", " ");
    
    // Remove multiple white spaces from HTML
    HTMLCode = Regex.Replace(HTMLCode, "\\s+", " ");
    
    // Remove any JavaScript
    HTMLCode = Regex.Replace(HTMLCode, "<script.*?</script>", ""
    , RegexOptions.IgnoreCase | RegexOptions.Singleline);
    
    // Replace special characters like &, <, >, " etc.
    StringBuilder sbHTML = new StringBuilder(HTMLCode);
    // Note: There are many more special characters, these are just
    // most common. You can add new characters in this arrays if needed
    string[] OldWords = {"&nbsp;", "&amp;", "&quot;", "&lt;", 
    "&gt;", "&reg;", "&copy;", "&bull;", "&trade;"};
    string[] NewWords = {" ", "&", "\"", "<", ">", "Â®", "Â©", "â€¢", "â„¢"};
    for(int i = 0; i < OldWords.Length; i++)
    {
        sbHTML.Replace(OldWords[i], NewWords[i]);
    }
    
    // Check if there are line breaks (<br>) or paragraph (<p>)
    sbHTML.Replace("<br>", "!!br!!");
    //sbHTML.Replace("<br ", "!!br ");
    sbHTML.Replace("<p ", "!!p ");

    // Replace style-br's with a normal br
    string result = System.Text.RegularExpressions.Regex.Replace(sbHTML.ToString(), "<br style[^>]*>", "!!br!!");
    result = System.Text.RegularExpressions.Regex.Replace(result, "<div style[^>]*>", "!!br!!");
    
    // Remove all <tags>
    result = System.Text.RegularExpressions.Regex.Replace(
    result, "<[^>]*>", "");

    // Sub back in the <br>s and <p>s
    result = result.Replace("!!br!!", "<br>");
    //result = result.Replace("!!br", "<br");
    result = result.Replace("!!p", "<p");

    return result;
}
 
public static async Task<HttpResponseMessage> Run(HttpRequestMessage req, TraceWriter log)
{
    // Hooked in VSTS, whenever a feature's ring1_5 or ring3 field is updated.

    bool PRODUCTION = true;

    // Use the right Teams incoming hook
    List<string> TEAMS_HOOK_URIS = new List<string>();
    List<string> TEAMS_INTERNAL_HOOK_URIS = new List<string>();

    List<string> RING_1_5_3_HOOK_URIS = new List<string>();
    List<string> RING_4_HOOK_URIS = new List<string>();

    List<string> RING_1_5_3_INTERNAL_HOOK_URIS = new List<string>();
    List<string> RING_4_INTERNAL_HOOK_URIS = new List<string>();
    List<string> RING_1_INTERNAL_HOOK_URIS = new List<string>();
    List<string> RING_2_INTERNAL_HOOK_URIS = new List<string>();
    List<string> RING_3_9_INTERNAL_HOOK_URIS = new List<string>();

    string FIELD_NAMESPACE, VALIDATION_FIELD_NAMESPACE, PROJECT_NAME;
    string SUB_ID_1_5, SUB_ID_3, SUB_ID_4;
    string SUB_ID_1, SUB_ID_2, SUB_ID_3_9;

    bool FEEDBACK_DISABLED;
    if (PRODUCTION) {
        // Really really production channel
        RING_1_5_3_HOOK_URIS.Add("https://outlook.office.com/webhook/*sanitized*/IncomingWebhook/*sanitized*/*sanitized*");
        // Really really Skype production channel
        RING_1_5_3_HOOK_URIS.Add("https://outlook.office.com/webhook/*sanitized*/IncomingWebhook/*sanitized*/*sanitized*");

        // TAP100 R4 channel
        RING_4_HOOK_URIS.Add("https://outlook.office.com/webhook/*sanitized*/IncomingWebhook/*sanitized*/*sanitized*");
        // Skype R4 channel
        RING_4_HOOK_URIS.Add("https://outlook.office.com/webhook/*sanitized*/IncomingWebhook/*sanitized*/*sanitized*");

        // New webhooks for internal MS users; they get a different card with more fields, and more rings to keep track of
        RING_1_5_3_INTERNAL_HOOK_URIS.Add("https://outlook.office.com/webhook/*sanitized*/IncomingWebhook/*sanitized*/*sanitized*");
        RING_4_INTERNAL_HOOK_URIS.Add("https://outlook.office.com/webhook/*sanitized*/IncomingWebhook/*sanitized*/*sanitized*");
        RING_1_INTERNAL_HOOK_URIS.Add("https://outlook.office.com/webhook/*sanitized*/IncomingWebhook/*sanitized*/*sanitized*");
        RING_2_INTERNAL_HOOK_URIS.Add("https://outlook.office.com/webhook/*sanitized*/IncomingWebhook/*sanitized*/*sanitized*");
        RING_3_9_INTERNAL_HOOK_URIS.Add("https://outlook.office.com/webhook/*sanitized*/IncomingWebhook/*sanitized*/*sanitized*");

        FIELD_NAMESPACE = "MicrosoftTeamsCMMI";
        VALIDATION_FIELD_NAMESPACE = "MicrosoftTeamsCMMI-Copy";
        PROJECT_NAME = "MSTeams";
        SUB_ID_1_5 = "cefb2308-2c11-4b88-a9d8-4db94270b756";
        SUB_ID_3 = "aeff807e-117e-4790-a353-334baf515fd9";
        SUB_ID_4 = "329348a5-0060-4640-90b6-94e8ef048118";

        // Don't know these yet
        SUB_ID_1 = "2b19dac9-b26e-45df-9a5b-9af21242bb9b";
        SUB_ID_2 = "b338c37c-84fb-4b84-a9f0-5d248901ad64";
        SUB_ID_3_9 = "a21d70e3-dd7f-4cd3-9088-fe70df8e8ada";
        FEEDBACK_DISABLED = true;
    } else {
        // Hook URI for TAP Feature Announcements Test channel
        //RING_1_5_3_HOOK_URIS.Add("https://outlook.office.com/webhook/37317ed8-68c1-4564-82bb-d2acc4c6b2b4@72f988bf-86f1-41af-91ab-2d7cd011db47/IncomingWebhook/b4073f0edef34b899ea708a4b6659978/512d26c9-aeed-4dbd-a16f-398bcf0ec3fe");
        // Hook URI for VanceFridge channel
        RING_1_5_3_HOOK_URIS.Add("https://outlook.office.com/webhook/15f69bfd-b260-478c-af37-db567141623a@240a3177-7696-41df-a4ea-0d1b0999fb38/IncomingWebhook/a79736f03e834b05b66a805574d20cc3/2c46ba10-5141-459c-a866-d47436de0cba");
        
        RING_4_HOOK_URIS.Add("https://outlook.office.com/webhook/15f69bfd-b260-478c-af37-db567141623a@240a3177-7696-41df-a4ea-0d1b0999fb38/IncomingWebhook/5a2da2b0ece94c75aefe55198eb432c3/2c46ba10-5141-459c-a866-d47436de0cba");

        FIELD_NAMESPACE = "Custom";
        VALIDATION_FIELD_NAMESPACE = "Custom";
        //PROJECT_NAME = "MSTeams";
        PROJECT_NAME = "MyFirstProject";
        
        SUB_ID_1_5 = "8303c1eb-d697-44cb-84ae-16d20b0ac207";
        SUB_ID_3 =   "c5b45aef-f2c7-45fb-aafb-4c6930f9e962";
        SUB_ID_4 = "329348a5-0060-4640-90b6-94e8ef048118";

        SUB_ID_1 = "";
        SUB_ID_2 = "";
        SUB_ID_3_9 = "";

        FEEDBACK_DISABLED = false;
    }

    log.Info("Webhook was triggered!");
    HttpContent requestContent = req.Content;
    string jsonContent = requestContent.ReadAsStringAsync().Result;

    log.Info(jsonContent);

    dynamic jsonObj = JObject.Parse(jsonContent);
    string feature = jsonObj.message.text;

    int featureId = jsonObj.resource.workItemId;
    string subscriptionId = jsonObj.subscriptionId;
    if ((subscriptionId != SUB_ID_1_5) && (subscriptionId != SUB_ID_3) && (subscriptionId != SUB_ID_4) && (subscriptionId != SUB_ID_1) && (subscriptionId != SUB_ID_2) && (subscriptionId != SUB_ID_3_9)) {
        log.Info("subscriptionId is " + subscriptionId + "; this sub not supported yet");
        return req.CreateResponse(HttpStatusCode.OK);
    }

    string featureTitle = jsonObj.resource.revision.fields["System.Title"];
    log.Info("Got title");    
    string featureDescription = "";
    try {
        featureDescription = jsonObj.resource.revision.fields["System.Description"].ToString();
    } catch (Exception e) {
        featureDescription = "No description given";
    }

    log.Info("Description length is " + featureDescription.Length);
    if (featureDescription.Length > 1500) {
        featureDescription = HtmlToText(featureDescription);
    }
    log.Info("Description length is now " + featureDescription.Length + " after converting to plaintext");

    // Truncate if it's still too long
    if (featureDescription.Length > 1500) {
        featureDescription = featureDescription.Substring(0, 1500) + "...";
    }
    log.Info("Futher truncated it");
    
    
    string featureAuthor = jsonObj.resource.revision.fields["System.ChangedBy"];

    // Feature's target date
    string featureTargetDatetime = "01/01/9999";
    string featureTargetDate = featureTargetDatetime;
    try {
        featureTargetDatetime = jsonObj.resource.revision.fields[FIELD_NAMESPACE + ".Ring4TargetDate"];
        featureTargetDate = featureTargetDatetime.Substring(0, 10);
    } catch (Exception e) { }

    // Feature's "TAP validation required?" field

    string featureValidationRequired = "Not specified";
    try {
        featureValidationRequired = jsonObj.resource.revision.fields[VALIDATION_FIELD_NAMESPACE + ".ERP_TAP_ValidationRequired"].ToString();
        if (featureValidationRequired == "")
        {
            featureValidationRequired = "Not specified";
        }
    } catch (Exception e) { }
    
    string featureUrl = jsonObj.resource._links.parent.href;
    string featureHumanUrl = "https://domoreexp.visualstudio.com/MSTeams/_workitems/edit/" + featureId;

    // Fields for the internal notifications
    // Testing
    string testPlan = "";
    try {
        testPlan = jsonObj.resource.revision.fields["MicrosoftTeamsCMMI-Copy.ERP_TestPlan"];
    } catch (Exception e) { };

    string testingScenario = "";
    try {
        testingScenario = jsonObj.resource.revision.fields["MicrosoftTeamsCMMI-Copy.ERP_Testing_Scenario"];
    } catch (Exception e) { };

    string testingManualTests = "";
    try {
        testingManualTests = jsonObj.resource.revision.fields["MicrosoftTeamsCMMI-Copy.ERP_Testing_ManualTests"];
    } catch (Exception e) { };

    string testingFullTestPass = "";
    try {
        testingFullTestPass = jsonObj.resource.revision.fields["MicrosoftTeamsCMMI-Copy.ERP_Testing_FullTestPass"];
    } catch (Exception e) { };

    string testingOffshore = "";
    try {
        testingOffshore = jsonObj.resource.revision.fields["MicrosoftTeamsCMMI-Copy.ERP_Testing_Offshore"];
    } catch (Exception e) { };

    // Support and Tech Readiness
    string customerReadinessImpact = "";
    try {
        customerReadinessImpact = jsonObj.resource.revision.fields["MicrosoftTeamsCMMI-New.CustomerReadinessImpact"];
    } catch (Exception e) { };
    
    string techReadiness = "";
    try {
        techReadiness = jsonObj.resource.revision.fields["MicrosoftTeamsCMMI.TechReadiness"];
    } catch (Exception e) { };

    string techReadinessO365 = "";
    try {
        techReadinessO365 = jsonObj.resource.revision.fields["MicrosoftTeamsCMMI-Copy.ERP_TR_O365"];
    } catch (Exception e) { };

    // TAP stuff
    // featureValidationRequired is the first TAP field
    string tapSignoff = "";
    try {
        tapSignoff = jsonObj.resource.revision.fields["MicrosoftTeamsCMMI-Copy.ERP_TAPChecklist_TAPSignoff"];
    } catch (Exception e) { };

    // Figure out which rings got enabled in this revision
    string ringValue = "";
    string availableRing = "Ring";
    int availableRingCount = 0;

    // Looking for specific ring changes in the payload.
    string[] rings = new[] {"1", "2", "1_5", "3", "3_9", "4"};

    string ringFieldName;

    foreach (string ring in rings)
    {
        log.Info("Checking out ring " + ring);
        try {
            // Somehow, the test VSTS has inconsistent namespaces. This'll have to do here
            if (!PRODUCTION)
            {
                if (ring == "1_5")
                {
                    ringFieldName = "Custom.Ring1_5";
                } else {
                    ringFieldName = "AgileRings.Ring" + ring;
                }
            }
            else {
                if (ring == "3_9")
                {
                    ringFieldName = VALIDATION_FIELD_NAMESPACE + ".Ring" + ring;
                } else {
                    ringFieldName = FIELD_NAMESPACE + ".Ring" + ring;
                }
                
            }

            log.Info(ringFieldName);
            log.Info(jsonObj.resource.fields[ringFieldName].ToString());
            
            ringValue = jsonObj.resource.fields[ringFieldName].newValue;
            log.Info("ringValue set");
            if ((ringValue == "Code available + Enabled") || (ringValue == "Code unavailable + Enabled")) {
                availableRing += " " + ring + " +";
                availableRingCount += 1;
            }
            
            log.Info(ringValue);

        } catch (Exception e) {}
    }

    // Make it more readable
    availableRing = availableRing.Replace("1_5", "1.5");
    availableRing = availableRing.Replace("3_9", "3.9");
    

    if (availableRingCount > 1) {
        if (availableRing.Contains("1 ")) {
            if (subscriptionId != SUB_ID_1) {
                log.Info("This is a duplicate notification. Skipping this one");
                return req.CreateResponse(HttpStatusCode.OK);
            }
        } else if (availableRing.Contains("1.5")) {
            if (subscriptionId != SUB_ID_1_5) {
                log.Info("This is a duplicate notification. Skipping this one");
                return req.CreateResponse(HttpStatusCode.OK);
            }
        } else if (availableRing.Contains("2")) {
            if (subscriptionId != SUB_ID_2) {
                log.Info("This is a duplicate notification. Skipping this one");
                return req.CreateResponse(HttpStatusCode.OK);
            }
        } else if (availableRing.Contains("3 ")) {
            if (subscriptionId != SUB_ID_3) {
                log.Info("This is a duplicate notification. Skipping this one");
                return req.CreateResponse(HttpStatusCode.OK);
            }
        } else if (availableRing.Contains("3.9")) {
            if (subscriptionId != SUB_ID_3_9) {
                log.Info("This is a duplicate notification. Skipping this one");
                return req.CreateResponse(HttpStatusCode.OK);
            }
        } else if (availableRing.Contains("4")) {
            if (subscriptionId != SUB_ID_4) {
                log.Info("This is a duplicate notification. Skipping this one");
                return req.CreateResponse(HttpStatusCode.OK);
            }
        }
    }

    // Check for whether it's enabled. If not, return and do nothing
    if (availableRingCount < 1) {
        log.Info("No rings were made available in this change. Ignoring");
        return req.CreateResponse(HttpStatusCode.OK);
    }

    log.Info("available ring is " + availableRing);

    // Send the notification to the right destination, depending on which rings are newly enabled
    if ((availableRing.Contains("1.5") || (availableRing.Contains("3 ")))) {
        TEAMS_HOOK_URIS.AddRange(RING_1_5_3_HOOK_URIS);
        TEAMS_INTERNAL_HOOK_URIS.AddRange(RING_1_5_3_INTERNAL_HOOK_URIS);
    }

    if (availableRing.Contains("4")) {
        TEAMS_HOOK_URIS.AddRange(RING_4_HOOK_URIS);
        TEAMS_INTERNAL_HOOK_URIS.AddRange(RING_4_INTERNAL_HOOK_URIS);
    }

    if (availableRing.Contains("1 ")) {
        TEAMS_INTERNAL_HOOK_URIS.AddRange(RING_1_INTERNAL_HOOK_URIS);
    }

    if (availableRing.Contains("2")) {
        TEAMS_INTERNAL_HOOK_URIS.AddRange(RING_2_INTERNAL_HOOK_URIS);
    }

    if (availableRing.Contains("3.9")) {
        TEAMS_INTERNAL_HOOK_URIS.AddRange(RING_3_9_INTERNAL_HOOK_URIS);
    }

    // Remove trailing spaces.... after we check for "1 " and "3 "
    availableRing = availableRing.Trim( new Char[] { ' ', '+' } );

    /*
        Get the enablement status of rings 1.5, 3, and 4.
    */

    bool ring1_5, ring3, ring4, ring1, ring2, ring3_9;
    ring1_5 = false;
    ring3 = false;
    ring4 = false;

    ring1 = false;
    ring2 = false;
    ring3_9 = false;

    ringFieldName = FIELD_NAMESPACE;
    try {
        if (!PRODUCTION) {
            ringFieldName = "Custom.Ring1_5";
        } else {
            ringFieldName = FIELD_NAMESPACE + ".Ring1_5";
        }
        string fieldValue = jsonObj.resource.revision.fields[ringFieldName];
        ring1_5 = ((fieldValue == "Code available + Enabled") || (fieldValue == "Code unavailable + Enabled"));
    } catch (Exception e) { }

    try {
        if (!PRODUCTION) {
            ringFieldName = "AgileRings.Ring3";
        } else {
            ringFieldName = FIELD_NAMESPACE + ".Ring3";
        }
        string fieldValue = jsonObj.resource.revision.fields[ringFieldName];
        ring3 = ((fieldValue == "Code available + Enabled") || (fieldValue == "Code unavailable + Enabled"));
    } catch (Exception e) { }

    try {
        if (!PRODUCTION) {
            ringFieldName = "AgileRings.Ring4";
        } else {
            ringFieldName = FIELD_NAMESPACE + ".Ring4";
        }
        string fieldValue = jsonObj.resource.revision.fields[ringFieldName];
        ring4 = ((fieldValue == "Code available + Enabled") || (fieldValue == "Code unavailable + Enabled"));
    } catch (Exception e) { }

    try {
        ringFieldName = FIELD_NAMESPACE + ".Ring1";
        string fieldValue = jsonObj.resource.revision.fields[ringFieldName];
        ring1 = ((fieldValue == "Code available + Enabled") || (fieldValue == "Code unavailable + Enabled"));
    } catch (Exception e) { }

    try {
        ringFieldName = FIELD_NAMESPACE + ".Ring2";
        string fieldValue = jsonObj.resource.revision.fields[ringFieldName];
        ring2 = ((fieldValue == "Code available + Enabled") || (fieldValue == "Code unavailable + Enabled"));
    } catch (Exception e) { }

    try {
        // 3.9 is "MicrosoftTeamsCMMI-Copy.Ring3_9"
        ringFieldName = VALIDATION_FIELD_NAMESPACE + ".Ring3_9";
        string fieldValue = jsonObj.resource.revision.fields[ringFieldName];
        ring3_9 = ((fieldValue == "Code available + Enabled") || (fieldValue == "Code unavailable + Enabled"));
    } catch (Exception e) { }



    log.Info(ring1 + " " + ring1_5 + " " + ring2 +  " " + ring3 + " " + ring3_9 + " " + ring4);

    
    string tapSafeAvailableRing = availableRing.Replace("+ 3.9", "").Replace("+ 2", "").Replace("1 + ", " ");

    // Generate the card that will show up in Teams
    dynamic betterCard = new {
        summary = "Feature now available for " + tapSafeAvailableRing,
        title = "Feature now available for " + tapSafeAvailableRing,
        sections = new object [] {
            new {
                activityTitle = featureTitle,
                activitySubtitle = featureDescription,
                activityImage = "https://i.imgur.com/xqG1HMv.png",
                facts = new [] {
                    new {
                        name = "Author",
                        value = featureAuthor
                    },
                    new {
                        name = "Feature expected to GA",
                        value = featureTargetDate
                    },
                    new {
                        name = "Ring 1.5",
                        value = ring1_5 ? "Enabled" : "Not enabled"
                    },
                    new {
                        name = "Ring 3",
                        value = ring3 ? "Enabled" : "Not enabled"
                    },
                    new {
                        name = "Ring 4",
                        value = ring4 ? "Enabled" : "Not enbaled"
                    },
                    new {
                        name = "TAP Validation Required?",
                        value = featureValidationRequired
                    }/*,
                    new {
                        name = "Pass/Fail",
                        value = "0/0"
                    }*/
                }
            }
        },
    };

    dynamic internalCard = new {
        summary = "Feature now available for " + availableRing,
        title = "Feature now available for " + availableRing,
        sections = new object [] {
            new {
                activityTitle = "[#" + featureId + ": " + featureTitle + "](" + featureHumanUrl + ")",
                activitySubtitle = featureDescription,
                activityImage = "https://i.imgur.com/xqG1HMv.png",
                facts = new [] {
                    new {
                        name = "Feature expected to GA",
                        value = featureTargetDate
                    },
                    new {
                        name = "Ring 1",
                        value = ring1 ? "Enabled" : "Not enabled"
                    },
                    new {
                        name = "Ring 1.5",
                        value = ring1_5 ? "Enabled" : "Not enabled"
                    },
                    new {
                        name = "Ring 2",
                        value = ring2 ? "Enabled" : "Not enabled"
                    },
                    new {
                        name = "Ring 3",
                        value = ring3 ? "Enabled" : "Not enabled"
                    },
                    new {
                        name = "Ring 3.9",
                        value = ring3_9 ? "Enabled" : "Not enabled"
                    },
                    new {
                        name = "Ring 4",
                        value = ring4 ? "Enabled" : "Not enbaled"
                    },
                }
            },
            new {
                activityTitle = "Testing and Automation",
                facts = new [] {
                    new {
                        name = "Test Plan",
                        value = testPlan
                    },
                    new {
                        name = "Scenario/Unit Tests",
                        value = testingScenario
                    },
                    new {
                        name = "Manual Tests",
                        value = testingManualTests
                    },
                    new {
                        name = "Manual Tests full test pass is scheduled on",
                        value = testingFullTestPass
                    },
                    new {
                        name = "Manual Tests - offshore vendor test pass",
                        value = testingOffshore
                    }
                }
            },
            new {
                activityTitle = "Support and Tech Readiness",
                facts = new [] {
                    new {
                        name = "Customer Readiness Impact",
                        value = customerReadinessImpact
                    },
                    new {
                        name = "TR Package",
                        value = techReadiness
                    },
                    new {
                        name = "O365 message center required before R4?",
                        value = techReadinessO365
                    }
                }
            },
            new {
                activityTitle = "TAP",
                facts = new [] {
                    new {
                        name = "Do you need TAP validation?",
                        value = featureValidationRequired
                    },
                    new {
                        name = "TAP Signoff",
                        value = tapSignoff
                    }
                }
            }
        }
    };

    JObject betterCardJObject = JObject.FromObject(betterCard);
    string betterCardObjectString = betterCardJObject.ToString();

    JObject internalCardJObject = JObject.FromObject(internalCard);


    string betterCardJson = betterCardJObject.ToString();
    // This does some double-replacement, converting type to @@type.
    betterCardJson = betterCardJson.Replace("type", "@type");
    // ...and here's how I fixed that:
    betterCardJson = betterCardJson.Replace("@@type", "@type");

    string internalCardJson = internalCardJObject.ToString();

    foreach (string teams_hook_uri in TEAMS_HOOK_URIS)
    {
        using (var client = new HttpClient())
        {
            var response = await client.PostAsync(
                teams_hook_uri,
                new StringContent(betterCardJson, System.Text.Encoding.UTF8, "application/json")
            );
        }
    }

    // Send internal card to internal folks
    foreach (string internal_hook_uri in TEAMS_INTERNAL_HOOK_URIS)
    {
        log.Info("Sending internalCardJson to " + internal_hook_uri);
        using (var client = new HttpClient())
        {
            var response = await client.PostAsync(
                internal_hook_uri,
                new StringContent(internalCardJson, System.Text.Encoding.UTF8, "application/json")
            );
        }
    }

    return req.CreateResponse(HttpStatusCode.OK);
}
