function manageGitlabData() {
    console.info("Starting execution of manageGitlabData function.");
    const startDate = PHM.Properties.getProp('GITLAB_START_DATE');
    const endDate = PHM.Properties.getProp('GITLAB_END_DATE');
    // Start processing the GitLab data
    fetchGitlabData(startDate, endDate);
}

function resetGitlabProcess() {
    console.info("Starting to reset GitLab process...");
    // Get the configuration sheet and the date span for the changelog
    const configSheet = PHM.Spreadsheet.getSheetByName(CONFIG_SHEET_NAME);
    const datespan = configSheet.getRange(CHGLOG_DATESPAN_RANGE).getValue();
    console.info(`Date span for changelog set to ${datespan} days.`);

    // Calculate the start and end dates for the changelog based on the date span
    const endDate = new Date(); // Current date
    const startDate = new Date(endDate.getTime() - (datespan * 24 * 60 * 60 * 1000)); // Calculate start date

    // Store initial values in script properties for future reference
    PHM.Properties.setProp('GITLAB_START_DATE', startDate.toISOString());
    PHM.Properties.setProp('GITLAB_END_DATE', endDate.toISOString());
    console.info(`Start and End dates stored in script properties.`);

    PHM.Properties.setProp('GITLAB_FILE_ID', PHM.Utilities.createAndStoreJSONFile('gitlab.json', JSON.stringify([])).getId());

    PHM.Properties.setProp('GITLAB_CURRENT_CURSOR', '');
    PHM.Properties.setProp('GITLAB_HAS_NEXT_PAGE', true);
    console.info("GitLab process reset completed.");
}

function writeGitlabWIPData() {
    console.info("Writing GitLab WIP data to the sheet...");

    // Open the JSON file to retrieve raw data
    const fileId = PHM.Properties.getProp('GITLAB_FILE_ID');
    const file = PHM.Utilities.openJSONFile(fileId);
    if (!file) {
        console.error('Failed to open JSON file.');
        return;
    }

    const fileDataString = file.getBlob().getDataAsString();
    let gitlabData;
    try {
        gitlabData = JSON.parse(fileDataString);
    } catch (error) {
        console.error('Failed to parse JSON data:', error);
        return;
    }

    const wipSheet = PHM.Spreadsheet.getSheetByName(GITLAB_WIP_SHEET_NAME);
    const rows = [];

    gitlabData.forEach(item => {
        const mergeId = item.id.split('/').pop(); // Extract numeric Merge ID
        const projectId = item.project.id.split('/').pop(); // Extract numeric Project ID
        const squad = item.squad || ''; // Squad
        const mrDate = PHM.DateUtils.formatDate(item.createdAt, true); // Format MR Date
        const totalApprovals = item.approvedBy.nodes.length; // Count total approvals
        const totalComments = item.commenters.nodes.length; // Count total comments
        const mergeDate = item.mergedAt ? PHM.DateUtils.formatDate(item.mergedAt, true) : ''; // Format Merge Date
        const leadTime = item.mergedAt ? PHM.DateUtils.calculateDuration(item.createdAt, item.mergedAt) : ''; // Calculate Lead Time
        const title = item.title; // Title
        const description = item.description; // Description
        const state = item.state; // State
        const authorName = item.author.name; // Author Name
        const mergedByName = item.mergeUser ? item.mergeUser.name : ''; // Merged By Name
        const webUrl = item.webUrl; // Web URL

        // Push the row data to the array
        rows.push([
            mergeId,
            projectId,
            squad, // Squad (to be filled if available)
            mrDate,
            totalApprovals,
            totalComments,
            mergeDate,
            leadTime,
            title,
            description,
            state,
            authorName,
            mergedByName,
            webUrl
        ]);
    });

    // Write the data to the sheet starting from row 2
    if (rows.length > 0) {

        wipSheet.getRange("A2:N").clearContent();
        wipSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
        console.info(`Successfully wrote ${rows.length} rows to the sheet.`);
    } else {
        console.info("No data to write to the sheet.");
    }
}

function writeGitlabChangelogData() {
    console.info("Writing GitLab Changelog data to the sheet...");

    // Open the JSON file to retrieve raw data
    const fileId = PHM.Properties.getProp('GITLAB_FILE_ID');
    const file = PHM.Utilities.openJSONFile(fileId);
    if (!file) {
        console.error('Failed to open JSON file.');
        return;
    }

    const fileDataString = file.getBlob().getDataAsString();
    let gitlabData;
    try {
        gitlabData = JSON.parse(fileDataString);
    } catch (error) {
        console.error('Failed to parse JSON data:', error);
        return;
    }

    const changelogSheet = PHM.Spreadsheet.getSheetByName(GITLAB_CHGLOG_SHEET_NAME);
    const rows = [];

    gitlabData.forEach(item => {
        const mergeId = 'gitlab-' + item.id.split('/').pop(); // Extract numeric Merge ID
        const projectName = item.project.name; // Project Name
        const squad = item.squad || ''; // Squad
        const createdAt = new Date(item.createdAt); // Date of MR creation

        // Action for MR creation
        rows.push([
            mergeId,
            PHM.DateUtils.formatDate(createdAt, true), // Date of MR creation
            'gitlab', // Tool
            projectName,
            squad,
            item.author.name, // Author
            'mr criado', // Action
            `criação de ${item.title} - ${mergeId}` // Detail
        ]);

        // Actions for approvals
        item.approvedBy.nodes.forEach(approval => {
            rows.push([
                mergeId,
                PHM.DateUtils.formatDate(createdAt, true), // Date of approval
                'gitlab', // Tool
                projectName,
                squad,
                approval.name, // Approver's name
                'mr aprovado', // Action
                `aprovação de ${item.title} - ${mergeId}` // Detail
            ]);
        });

        // Actions for comments
        item.commenters.nodes.forEach(commenter => {
            rows.push([
                mergeId,
                PHM.DateUtils.formatDate(createdAt, true), // Date of comment
                'gitlab', // Tool
                projectName,
                squad,
                commenter.name, // Commenter's name
                'comentário', // Action
                `comentário em ${item.title} - ${mergeId}` // Detail
            ]);
        });

        // Action for MR merge
        if (item.mergedAt) {
            rows.push([
                mergeId,
                PHM.DateUtils.formatDate(item.mergedAt, true), // Date of merge
                'gitlab', // Tool
                projectName,
                squad,
                item.mergeUser ? item.mergeUser.name : '', // Merged by name
                'merge', // Action
                `merge de ${item.title} - ${mergeId}` // Detail
            ]);
        }
    });

    // Write the data to the sheet starting from row 2
    if (rows.length > 0) {
        changelogSheet.getRange("A2:H").clearContent();
        changelogSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
        console.info(`Successfully wrote ${rows.length} rows to the changelog sheet.`);
    } else {
        console.info("No data to write to the changelog sheet.");
    }
    return rows.length;
}


function writeGitlabProjectsData() {
    console.info("Writing GitLab Projects data to the sheet...");

    // Open the JSON file to retrieve raw data
    const fileId = PHM.Properties.getProp('GITLAB_FILE_ID');
    const file = PHM.Utilities.openJSONFile(fileId);
    if (!file) {
        console.error('Failed to open JSON file.');
        return;
    }

    const fileDataString = file.getBlob().getDataAsString();
    let gitlabData;
    try {
        gitlabData = JSON.parse(fileDataString);
    } catch (error) {
        console.error('Failed to parse JSON data:', error);
        return;
    }

    const startDate = new Date(PHM.Properties.getProp('GITLAB_START_DATE'));
    const endDate = new Date(PHM.Properties.getProp('GITLAB_END_DATE'));
    const projectStats = {};

    // Process each merge request to compile project data
    gitlabData.forEach(item => {
        const projectName = item.project.name;
        const projectId = item.project.id.split('/').pop(); // Extract numeric Project ID
        const createdAt = new Date(item.createdAt);
        const mergedAt = item.mergedAt ? new Date(item.mergedAt) : null;
        const squad = item.squad; // Access squad from JSON

        // Initialize project data if not already present
        if (!projectStats[projectId]) {
            projectStats[projectId] = {
                projectName: projectName,
                squad: squad, // Fill in squad if available
                merges: 0,
                latestMergeDate: null,
                totalLeadTime: 0,
                mergeCount: 0
            };
        }

        // Increment merge count
        projectStats[projectId].merges += 1;

        // Update latest merge date
        if (mergedAt && (!projectStats[projectId].latestMergeDate || mergedAt > projectStats[projectId].latestMergeDate)) {
            projectStats[projectId].latestMergeDate = mergedAt;
        }

        // Calculate lead time if merged
        if (mergedAt) {
            const leadTime = PHM.DateUtils.calculateDuration(createdAt, mergedAt);
            projectStats[projectId].totalLeadTime += leadTime;
            projectStats[projectId].mergeCount += 1; // Count for average calculation
        }
    });

    // Prepare data for writing to the sheet
    const rows = [];
    for (const projectId in projectStats) {
        const stats = projectStats[projectId];
        const avgLeadTime = stats.mergeCount > 0 ? (stats.totalLeadTime / stats.mergeCount) : 0;
        const mergesPerMonth = stats.merges / ((endDate - startDate) / (1000 * 60 * 60 * 24 * 30)); // Average merges per month

        rows.push([
            stats.projectName,
            projectId,
            stats.squad, // Fill in squad
            stats.merges,
            stats.latestMergeDate ? PHM.DateUtils.formatDate(stats.latestMergeDate, true) : '', // Latest Merge Date
            avgLeadTime, // Leadtime for changes (AVG)
            mergesPerMonth // Merges per month (avg)
        ]);
    }

    // Sort rows by number of merges, descending
    rows.sort((a, b) => b[3] - a[3]); // Sort by merges column

    // Limit to top 100 projects
    const topProjects = rows.slice(0, 500);

    // Write the data to the sheet starting from row 2
    const projectsSheet = PHM.Spreadsheet.getSheetByName(GITLAB_SHEET_NAME);
    projectsSheet.getRange("A2:G").clearContent();
    if (topProjects.length > 0) {
        projectsSheet.getRange(2, 1, topProjects.length, topProjects[0].length).setValues(topProjects);
        console.info(`Successfully wrote ${topProjects.length} rows to the projects sheet.`);
    } else {
        console.info("No data to write to the projects sheet.");
    }
}

function fetchGitlabData(startDate, endDate) {
    let hasNextPage = PHM.Properties.getProp('GITLAB_HAS_NEXT_PAGE'); // Ensure boolean
    let itemsPerCall = 100;
    const maxItemsPerExecution = 1500;
    let itemsFetched = 0;
    const fileID = PHM.Properties.getProp('GITLAB_FILE_ID');
    let file;
    let jsonFileData = [];

    try {
        file = PHM.Utilities.openJSONFile(fileID);
        if (file) {
            const fileContent = file.getBlob().getDataAsString();
            jsonFileData = JSON.parse(fileContent);
        }
    } catch (error) {
        console.warn(`Error opening JSON file: ${error.message}. Proceeding to create a new file.`);
    }

    let allMergeRequests = []; // Variable to store all merge requests

    while (hasNextPage && itemsFetched < maxItemsPerExecution) {
        let cursor = PHM.Properties.getProp('GITLAB_CURRENT_CURSOR');
        const afterPart = cursor !== "null" ? `after: "${cursor}",` : "";

        // Adjust itemsPerCall if it exceeds the remaining limit
        if (itemsFetched + itemsPerCall > maxItemsPerExecution) {
            itemsPerCall = maxItemsPerExecution - itemsFetched;
            console.info(`Adjusting itemsPerCall to: ${itemsPerCall} to respect maxItemsPerExecution`);
        }

        const query = `
            query {
                group(fullPath: "${GITLAB_GROUP_ID}") {
                    mergeRequests(
                        includeSubgroups: true, 
                        first: ${itemsPerCall}, 
                        ${afterPart}
                        createdAfter: "${new Date(startDate).toISOString()}", 
                        createdBefore: "${new Date(endDate).toISOString()}"
                    ) {
                        count
                        pageInfo {
                            endCursor
                            hasNextPage
                        }
                        nodes {
                            id
                            iid
                            title
                            project {
                                id
                                name
                            }
                            createdAt
                            author {
                                name
                                username
                            }
                            mergedAt
                            mergeUser {
                                name
                                username
                            }
                            approvedBy {
                                nodes {
                                    name
                                    username
                                }
                            }
                            commenters {
                                nodes {
                                    name
                                    username
                                }
                            }
                            state
                            webUrl
                        }
                    }
                }
            }
        `;

        const options = {
            method: 'post',
            contentType: 'application/json',
            headers: {
                'Authorization': `Bearer ${GITLAB_TOKEN}`,
                'PRIVATE-TOKEN': `${GITLAB_TOKEN}`
            },
            payload: JSON.stringify({ query: query }),
            muteHttpExceptions: true
        };

        const response = UrlFetchApp.fetch(`${GITLAB_API_URL}`, options);
        const responseText = response.getContentText();
        const data = JSON.parse(responseText);

        if (data.errors) {
            console.error('GraphQL errors:', data.errors);
            return;
        }

        if (!data.data?.group?.mergeRequests) {
            console.error('Unexpected data structure:', data);
            return;
        }

        const mergeRequests = data.data.group.mergeRequests.nodes;
        const pageInfo = data.data.group.mergeRequests.pageInfo;
        var total = data.data.group.mergeRequests.count;
        

        // Store the fetched data in the allMergeRequests array
        allMergeRequests = allMergeRequests.concat(mergeRequests);

        // Update cursor and pagination state
        if (pageInfo.hasNextPage) PHM.Properties.setProp('GITLAB_HAS_NEXT_PAGE', pageInfo.hasNextPage.toString());
        if (pageInfo.endCursor) PHM.Properties.setProp('GITLAB_CURRENT_CURSOR', pageInfo.endCursor.toString());

        itemsFetched += mergeRequests.length;
        hasNextPage = pageInfo.hasNextPage;

        //console.info(`Cursor updated to: ${pageInfo.endCursor}`);
        //console.info(`Has next page: ${hasNextPage}`);
        console.info(`Items fetched so far: ${itemsFetched}`);

        if (itemsFetched >= maxItemsPerExecution) {
            console.info(`Reached the maximum of ${maxItemsPerExecution} items per execution.`);
            break; // Exit the loop *before* checking the file size
        }
    }

    console.info(`Number of merge requests to retrieve: ${total}`);

    // Update the existing file with the new data outside the while loop
    if (allMergeRequests.length > 0) {
        console.info(`We are going to process merge request squad, username, etc... Total: ${allMergeRequests.length}, starting time: ${new Date()}`);
        // Add a squad field to each merge request node
        allMergeRequests.forEach(mergeRequest => {
            // console.log(`Processing merge request: ${mergeRequest.id}`);
            const squad = squadByGitlabProjectId(mergeRequest.project.id);
            if (squad) {
                // console.log(`Found squad for project ID: ${mergeRequest.project.id}, squad: ${squad}`);
                mergeRequest.squad = squad;
            } else {
                // console.log(`No squad found for project ID: ${mergeRequest.project.id}, using project ID as squad.`);
            }
            //process users according to user dictionary
            if (mergeRequest.mergeUser) mergeRequest.mergeUser.name = PHM.Utilities.updateUsername(mergeRequest.mergeUser.name, USER_DICTIONARY) || mergeRequest.mergeUser.name + "*"; // Merged By Name
            if (mergeRequest.author) mergeRequest.author.name = PHM.Utilities.updateUsername(mergeRequest.author.name, USER_DICTIONARY) || mergeRequest.author.name + "*"; // Author Name
            if (mergeRequest.approvedBy.nodes) {
                mergeRequest.approvedBy.nodes.forEach(approval => {
                    approval.name = PHM.Utilities.updateUsername(approval.name, USER_DICTIONARY) || approval.name + "*"; // Approved By Name
                });
            }
            if (mergeRequest.commenters.nodes) {
                mergeRequest.commenters.nodes.forEach(commenter => {
                    commenter.name = PHM.Utilities.updateUsername(commenter.name, USER_DICTIONARY) || commenter.name + "*"; // Commenter Name
                });
            }
        });

        console.log(`All merge requests processed. Total: ${allMergeRequests.length}, ending time: ${new Date()}`);

        jsonFileData = jsonFileData.concat(...allMergeRequests);

        try {
            file = PHM.Utilities.createAndStoreJSONFile('gitlab.json', JSON.stringify(jsonFileData, null, 2));
            PHM.Properties.setProp('GITLAB_FILE_ID', file.getId());
            // console.info(`Data stored on JSON file. New size: ${PHM.Utilities.countNodesInJSONFile(PHM.Properties.getProp('GITLAB_FILE_ID'))} nodes`);
        } catch (error) {
            console.error(`Failed to create and store new file: ${error.message}`);
        }
    }

    // Final count after all data is retrieved or maxItemsPerExecution is reached
    const updatedFileContent = file.getBlob().getDataAsString();
    const updatedJsonData = JSON.parse(updatedFileContent);
    console.info(`Total nodes in the file after update: ${updatedJsonData.length}`);

    if (!hasNextPage) {
        console.info("All data has been retrieved from GitLab. Proceeding to next stage.");
        nextScriptExecutionStage();
    }
}

// Move the function outside the fetchGitlabData function
function squadByGitlabProjectId(projectId) {
    const squadMapping = PHM.Spreadsheet.getRangeValues(GITLAB_PROJECT_IDS_SQUADS);

    // Verify if the data was loaded correctly
    if (!squadMapping || squadMapping.length === 0) {
        throw new Error("Gitlab squad mapping data is empty or could not be retrieved.");
    }

    // Create an object for efficient lookup
    const squadMap = squadMapping.reduce((acc, [id, squad]) => {
        acc[id] = squad;
        return acc;
    }, {});

    // Extract the numeric ID from the projectId URL
    const numericProjectId = projectId.replace(/[^0-9]/g, '').toString();

    // Return the squad from the map
    return squadMap[numericProjectId];
}
