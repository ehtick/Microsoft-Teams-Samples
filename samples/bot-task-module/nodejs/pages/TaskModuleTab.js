//- Copyright (c) Microsoft Corporation.
//- All rights reserved.
// Microsoft Teams app initialization
microsoftTeams.app.initialize().then(() => {
    const taskModuleButtons = document.getElementsByClassName("taskModuleButton");
    if (taskModuleButtons.length > 0) {
        fetch(`${window.location.origin}/getAppConfig`).then(response => response.json()).then(data => {

            // Configure your app ID and base URL
            const config = {
                MicrosoftAppID: data.MicrosoftAppId,
                BaseUrl: `${window.location.origin}`
            };

            const taskInfo = {
                title: "",
                size: "",
                url: "",
                card: "",
                fallbackUrl: "",
                completionBotId: config.MicrosoftAppID
            };

            const TaskModuleIds = {
                YouTube: "youtube",
                PowerApp: "powerapp",
                CustomForm: "customform",
                AdaptiveCard1: "adaptivecard1",
                AdaptiveCard2: "adaptivecard2"
            };

            const DeepLinkIds = {
                CustomForm: "customform.html",
            };

            // Titles for task modules
            const TaskModuleStrings = {
                YouTubeTitle: "Microsoft Ignite 2018 Vision Keynote",
                PowerAppTitle: "PowerApp: Asset Checkout",
                CustomFormTitle: "Custom Form",
                AdaptiveCardTitle: "Create a new job posting",
                AdaptiveCardKitchenSinkTitle: "Adaptive Card: Inputs",
                ActionSubmitResponseTitle: "Action.Submit Response",
                YouTubeName: "YouTube",
                PowerAppName: "PowerApp",
                CustomFormName: "Custom Form",
                AdaptiveCardSingleName: "Adaptive Card - Single",
                AdaptiveCardSequenceName: "Adaptive Card - Sequence"
            };

            // Sizes for task modules
            const TaskModuleSizes = {
                youtube: {
                    width: 1000,
                    height: 700
                },
                powerapp: {
                    width: 720,
                    height: 520
                },
                customform: {
                    width: 510,
                    height: 430
                },
                adaptivecard: {
                    width: 700,
                    height: 255
                }
            };

            function appRoot() {
                if (typeof window === "undefined") {
                    return config.BaseUrl;
                } else {
                    return `${window.location.protocol}//${window.location.host}`;
                }
            }

            // Initialize DeepLink
            const deepLink = document.getElementById("deeplink");
            deepLink.href = `https://teams.microsoft.com/l/task/${config.MicrosoftAppID}?url=${appRoot()}/${DeepLinkIds.CustomForm}&height=${TaskModuleSizes.customform.height}&width=${TaskModuleSizes.customform.width}&title=${TaskModuleStrings.CustomFormTitle}&completionBotId=${config.MicrosoftAppID}`;

            // Adaptive Card Template
            const adaptiveCardTemplate = {
                type: "AdaptiveCard",
                body: [
                    {
                        type: "TextBlock",
                        separator: true,
                        size: "Large",
                        weight: "Bolder",
                        text: "Enter basic information for this position:",
                        isSubtle: true,
                        wrap: true
                    },
                    {
                        type: "TextBlock",
                        separator: true,
                        text: "Title",
                        wrap: true
                    },
                    {
                        type: "Input.Text",
                        id: "jobTitle",
                        placeholder: "E.g. Senior PM"
                    },
                    {
                        type: "ColumnSet",
                        columns: [
                            {
                                type: "Column",
                                items: [
                                    {
                                        type: "TextBlock",
                                        text: "Level",
                                        wrap: true
                                    },
                                    {
                                        type: "Input.Number",
                                        id: "jobLevel",
                                        value: "7",
                                        placeholder: "Level in numbers min **1** and max **10**",
                                        min: 1,
                                        max: 10
                                    }
                                ],
                                width: 2
                            },
                            {
                                type: "Column",
                                items: [
                                    {
                                        type: "TextBlock",
                                        text: "Location"
                                    },
                                    {
                                        type: "Input.ChoiceSet",
                                        id: "jobLocation",
                                        value: "1",
                                        choices: [
                                            {
                                                title: "San Francisco",
                                                value: "1"
                                            },
                                            {
                                                title: "London",
                                                value: "2"
                                            },
                                            {
                                                title: "Singapore",
                                                value: "3"
                                            },
                                            {
                                                title: "Dubai",
                                                value: "3"
                                            },
                                            {
                                                title: "Frankfurt",
                                                value: "3"
                                            }
                                        ],
                                        isCompact: true
                                    }
                                ],
                                width: 2
                            }
                        ]
                    }
                ],
                actions: [
                    {
                        type: "Action.Submit",
                        id: "createPosting",
                        title: "Create posting",
                        data: {
                            command: "createPosting",
                            taskResponse: "{{responseType}}"
                        }
                    },
                    {
                        type: "Action.Submit",
                        id: "cancel",
                        title: "Cancel"
                    }
                ],
                version: "1.0"
            };

            // Add event listeners to buttons
            for (const btn of taskModuleButtons) {
                btn.addEventListener("click", function () {
                    taskInfo.url = `${appRoot()}/${this.id.toLowerCase()}.html`;

                    // Define default submitHandler()
                    let submitHandler = (err, result) => {
                        console.log(`Err: ${err}; Result: ${result}`);
                    };

                    switch (this.id.toLowerCase()) {
                        case TaskModuleIds.YouTube:
                            taskInfo.title = TaskModuleStrings.YouTubeTitle;
                            taskInfo.size = { height: TaskModuleSizes.youtube.height, width: TaskModuleSizes.youtube.width };
                            microsoftTeams.dialog.url.open(taskInfo, submitHandler);
                            break;

                        case TaskModuleIds.PowerApp:
                            taskInfo.title = TaskModuleStrings.PowerAppTitle;
                            taskInfo.size = { height: TaskModuleSizes.powerapp.height, width: TaskModuleSizes.powerapp.width };
                            microsoftTeams.dialog.url.open(taskInfo, submitHandler);
                            break;

                        case TaskModuleIds.CustomForm:
                            taskInfo.title = TaskModuleStrings.CustomFormTitle;
                            taskInfo.size = { height: TaskModuleSizes.customform.height, width: TaskModuleSizes.customform.width };

                            // SubmitHandler callback function
                            submitHandler = (err, result) => {
                                console.log(`Submit handler - err: ${err}`);
                                console.log(`Submit handler - result\rName: ${result.name}\rEmail: ${result.email}\rFavorite book: ${result.favoriteBook}`);
                            };

                            // Allows app to open a URL-based dialog.
                            microsoftTeams.dialog.url.open(taskInfo, submitHandler);
                            break;

                        case TaskModuleIds.AdaptiveCard1:
                            taskInfo.title = TaskModuleStrings.AdaptiveCardTitle;
                            taskInfo.size = { height: TaskModuleSizes.adaptivecard.height, width: TaskModuleSizes.adaptivecard.width };
                            taskInfo.card = adaptiveCardTemplate;

                            // SubmitHandler callback function
                            submitHandler = (err, result) => {
                                console.log(`Submit handler - err: ${err}`);
                                console.log(`Result = ${JSON.stringify(result)}\nError = ${JSON.stringify(err)}`);
                            };

                            // Dialogs (referred as task modules in TeamsJS v1.x) invoked from a tab
                            microsoftTeams.dialog.adaptiveCard.open(taskInfo, submitHandler);
                            break;

                        default:
                            console.log(`Unexpected button ID: ${this.id.toLowerCase()}`);
                            return;
                    }
                    console.log(`URL: ${taskInfo.url}`);
                });
            }
        });
    }
});