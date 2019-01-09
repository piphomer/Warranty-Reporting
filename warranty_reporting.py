#Python 3

from jira import JIRA
import csv
import sys
import os
import requests
from datetime import datetime as dt
import unicodedata
from unidecode import unidecode

#List of service desks to iterate through
service_desk_list = ['PS']

#Auth
jira_username = os.environ['JIRA_USERNAME']
jira_password = os.environ['JIRA_PASSWORD']



####################################################################################################


if __name__ == "__main__":

    run_date = dt.strftime(dt.now(), '%y%m%d')

    print(run_date)

    # Connect to BBOXX Jira server
    jira = JIRA('https://bboxxltd.atlassian.net', basic_auth=(jira_username, jira_password))

    
    #Get the tickets

    for desk_id in service_desk_list:

        #Define the header names of each column
        header_list = [
            "Ticket",
            "Issue Type",
            "Organization",
            "Reporter",
            "Product Type",
            "Failure",
            "Failure (Expanded)",
            "Quantity",
            "Sales Order(s)",
            "Assignee",
            "Description",
            "Status",
            "Resolution",
            "Created",
            "Updated",
            "Resolved",
            "Linked Issue"

        ]

        
        search_string = 'project = ' + desk_id + ' AND issuetype = "Warranty Claim"'

        print(search_string)

        # Retrieve all tickets so we can count them.
        # Note: we can only query 100 at a time so we will need to paginate in a later step
        all_tix = jira.search_issues(search_string)

        tix_count = all_tix.total # Count the tickets via the .total attribute

        page_qty = tix_count // 100 + 1 # Calculate how many pages of tickets there are

        #page_qty = 1  # Just run first page for test/debug

        print("Total number of tickets: ",tix_count)
        print("Number of pages: ", page_qty)

        output_list = []
        output_list_debug = []

        # Loop through the number of pages we need to gather all issues
        for page in range(page_qty):

            issue_list = []

            print("\r" + "Page ", page + 1, " of ", page_qty, end=' ')

            starting_issue = page * 100

            tix = jira.search_issues(search_string, startAt = starting_issue, maxResults= 100)

            for issue in tix:

                if issue.fields.issuetype.name == 'Warranty Claim':

                    print(issue.key)
                    
                    #organization
                    try:
                        organization = issue.raw['fields']['customfield_10700'][0]["name"]
                    except:
                        organization = "Unknown"

                    #Product type
                    try:
                        product_type = issue.raw['fields']['customfield_11407'][0]['value']
                    except:
                        product_type = 'Not specified'

                    #Make the dates Excel-readable
                    created = str(issue.fields.created)[:10] + " " + str(issue.fields.created)[11:19]
                    resolved = str(issue.fields.resolutiondate)[:10] + " " + str(issue.fields.resolutiondate)[11:19]
                    updated = str(issue.fields.updated)[:10] + " " + str(issue.fields.updated)[11:19]

                    #Assignee
                    try:
                        assignee = issue.fields.assignee.displayName
                    except:
                        assignee = "None"

                    #Resolution
                    try:
                        resolution = issue.fields.resolution.name
                    except:
                        resolution = "None"

                    #Status
                    try:
                        status = issue.fields.status.name
                    except:
                        status = "Unknown"

                    #Reporter
                    try:
                        reporter = issue.fields.reporter.displayName
                    except:
                        reporter = "Unknown"

                    #Affected Product Type
                    try:
                        product_type = issue.raw['fields']['customfield_11496']['value']
                    except:
                        product_type = "Unknown"

                    #Quantity of affected products
                    try:
                        quantity = issue.raw['fields']['customfield_11494']
                    except:
                        quantity = ""

                    #Affected sales orders
                    try:
                        sales_orders = issue.raw['fields']['customfield_11495']
                    except:
                        sales_orders = ""

                    #Failure description
                    try:
                        failure = issue.raw['fields']['customfield_11496']['child']['value']
                    except:
                        failure = "Unknown" 

                    #Failure (Expanded)
                    try:
                        failure_expanded = issue.raw['fields']['customfield_11497']
                    except:
                        failure_expanded = ""

                    #Linked issue
                    try:
                        linked_issue = issue.raw['fields']['issuelinks'][0]["inwardIssue"]["key"]
                    except:
                        linked_issue = "None"

                    

                    issue_list = [
                        issue.key,
                        issue.fields.issuetype.name,
                        organization,
                        reporter,
                        product_type,
                        failure,
                        failure_expanded,
                        quantity,
                        sales_orders,
                        assignee,
                        issue.fields.description,
                        status,
                        resolution,
                        created,
                        updated,
                        resolved,
                        linked_issue
                    ]

                    
                    output_list.append(issue_list)

                else:
                    pass


        #Write all metrics to Sharepoint
        fname = "{}_warranty_summary.csv".format(run_date)

        with open(fname, 'w', encoding='utf-8', newline='') as csvfile:
            
            print("Writing .csv file...")

            writer = csv.writer(csvfile, quoting=csv.QUOTE_NONNUMERIC)
            writer.writerow(header_list)
            writer.writerows(output_list)
