//tfx extension create

VSS.init({
    explicitNotifyLoaded: true,
    usePlatformStyles: true
});
        
VSS.require(["TFS/Dashboards/WidgetHelpers", "VSS/Authentication/Services", "TFS/WorkItemTracking/RestClient", "TFS/Work/RestClient"], function (WidgetHelpers, AuthService, WitApiClient, WorkApiClient) {
    WidgetHelpers.IncludeWidgetStyles();
    
    VSS.register("Hours-Counter", function () {   

        var HoursCounter = {

            CalculateHours: function(workItems) {

                var hours = 0;
                for (var i = 0; i < workItems.length; i++) {
                    hours += workItems[i].completedWorkToday;
                }
            
                return hours;
            },
        
            UpdateWidget: function(iterationName, workItems) {
            
                var totalHours = HoursCounter.CalculateHours(workItems);

                $("h2").html("You have logged");
                $(".big-count").html(totalHours);
                $(".footer").html("Hour" + ((totalHours === 1) ? "" : "s") + " today towards " + iterationName);

                $("div.widget a").on("click", function() {

                    if (totalHours > 0) {

                        VSS.getService(VSS.ServiceIds.Dialog).then(function(dialogService) {

                            var extensionContext = VSS.getExtensionContext();
                            var contributionId = extensionContext.publisherId + "." + extensionContext.extensionId + ".Hours-Modal";

                            dialogService.openDialog(contributionId, {
                                title: "My Hours Today",
                                contentText: "",
                                width: 600,
                                height: 250,
                                buttons: null
                            }).then(function(dialog) {

                                dialog.getContributionInstance("Hours-Modal").then(function (modalInstance) {
                                    modalInstance.loadHours(workItems);
                                });
                            });
                        });
                    }

                    return false;
                });
            },

            LoadWidget: function (widgetSettings) {

                return HoursCounter.GetCurrentIteration().then(function(currentIteration) {

                    // todo handle no iteration
                    if (!currentIteration) return;

                    var queryString = {
                        query:  "SELECT [System.Id] FROM WorkItems WHERE [System.WorkItemType] = 'Task' " +
                                "AND [System.IterationPath] = '" + currentIteration.path + "' " + 
                                "AND [Microsoft.VSTS.Scheduling.CompletedWork] <> '' AND [Microsoft.VSTS.Scheduling.CompletedWork] > 0 " +
                                "AND [System.ChangedDate] = @Today AND [System.AssignedTo] = @Me"
                    };
                    
                    return WitApiClient.getClient().queryByWiql(queryString).then(function (wiqlResult) {

                        var workItems = wiqlResult.workItems; 
                        if (workItems.length === 0) {
                            HoursCounter.UpdateWidget(currentIteration.name, workItems);
                        }
                                            
                        HoursCounter.AddCompletedHoursToWorkItems(workItems).then(function(processedWorkItems) {

                            HoursCounter.UpdateWidget(currentIteration.name, processedWorkItems);
                        });
                        
                        // Use the widget helper and return success as Widget Status
                        return WidgetHelpers.WidgetStatusHelper.Success();
                    }, function (error) {
                        // Use the widget helper and return failure as Widget Status
                        return WidgetHelpers.WidgetStatusHelper.Failure(error.message);
                    });                        
                });
            },

            AddCompletedHoursToWorkItems: function(workItems) {

                var deferred = $.Deferred();
                var tfsWebContext = VSS.getWebContext();
                var processedWorkItems = [];
                for (var i = 0; i < workItems.length; i++) {
                    
                    WitApiClient.getClient().getRevisions(workItems[i].id).then(function(revisionsResult) {
                        
                        if (revisionsResult.length == 0) return;
                        
                        var username = tfsWebContext.user.name + " <" + tfsWebContext.user.uniqueName + ">";
                        processedWorkItems.push({
                            id: revisionsResult[0].id,
                            completedWorkToday: HoursCounter.CalculateRevisionsHours(revisionsResult, username)
                        });

                        if (workItems.length === processedWorkItems.length) {
                            deferred.resolve(processedWorkItems);
                        }
                    });
                }

                return deferred.promise();
            },

            CalculateRevisionsHours: function(revisions, username) {
                
                var completedWork = 0;
                var previousCompletedWork = 0;

                for (var i = 0; i < revisions.length; i++) {
                    var revision = revisions[i];

                    if (i !== 0) {
                        var previousRevision = revisions[i - 1];
                        if (previousRevision.fields.hasOwnProperty("Microsoft.VSTS.Scheduling.CompletedWork")) {
                            previousCompletedWork = +previousRevision.fields["Microsoft.VSTS.Scheduling.CompletedWork"];
                            
                        }
                    };

                    if (revision.fields.hasOwnProperty("Microsoft.VSTS.Scheduling.CompletedWork") && new Date().setHours(0,0,0,0) === new Date(revision.fields["System.ChangedDate"]).setHours(0,0,0,0) && revision.fields["System.ChangedBy"] === username) {
                        completedWork += +revision.fields["Microsoft.VSTS.Scheduling.CompletedWork"] - previousCompletedWork;
                        
                    }
                }

                return completedWork;
            },

            GetCurrentIteration: function() {

                var deferred = $.Deferred();
                var tfsWebContext = VSS.getWebContext();

                $.ajax({
                    url: tfsWebContext.collection.relativeUri + tfsWebContext.project.name + "/" + tfsWebContext.team.name + "/_apis/work/TeamSettings/Iterations?$timeframe=current&api-version=2.0-preview.1",
                    contentType: 'application/json',
                    success: function(iterationData) {

                        var currentIteration = (iterationData.count === 1) ? iterationData.value[0] : undefined;
                        deferred.resolve(currentIteration);
                    }
                });

                return deferred.promise();
            }
        }

        return {
            load: function (widgetSettings) {

                return VSS.getAccessToken().then(function(token) {
                    var authHeader = AuthService.authTokenManager.getAuthorizationHeader(token);
                    $.ajaxSetup({
                        headers: {
                            Authorization: authHeader
                        }
                    });

                    return HoursCounter.LoadWidget(widgetSettings);
                });
            }
        }
    });

    VSS.notifyLoadSucceeded();
});