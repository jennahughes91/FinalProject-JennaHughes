# FinalProject-JennaHughes

## Context, User, and Problem

The target user is anyone who manages a backlog of development. The workflow that is being improved is the prioritization of the work captured in a development backlog. This problem matters because it is critical to always maintain an accurate, prioritized backlog of work so that you can ensure that you're always using your limited resources to work on what has the most impact to your business. What is in the backlog can routinely change so backlog re-prioritization is a workflow that may need to be performed on a near-constant basis for what can be dozens, to hundreds, or more of work items. 

## Solution and Design

I built an application that allows you to upload your excel spreadsheet backlog, give weights to business areas and product teams to generate a priority weight to each task and then rank them in ascending order. You can change the weights as many times as you like, save weights as a favorite to use again later, and then export a new excel file with those rankings. The main GenAI design choices were to use reasoning for the ranking system and few-shot prompting for the criteria, rule, and examples.

## Evaluation and Results

I compared against doing the workflow in a manual manner of performing the backlog prioritization in Excel. I ran 8 tests: missing columns, missing values, bad priority values, bad effort values, no item ID column, only one backlog item, an empty file, and column names that don’t match what is required. My evaluation showed that the test cases all had expected results in the app and they were more accurate and quicker than for my manual baseline. This was particularly important for bad priority values and unrecognized column names since those are most likely to appear in real backlogs. A human should still review the results of the app-generated prioritization as it is meant to be used as a helpful starting point not the end product. 

## Artifact Snapshot

https://youtu.be/4MjX0C7wQ2E 

## Setup and Usage Instructions

An excel XLS or XLSX file is required to be uploaded to the application with the below column names and data present within the rows for those columns. If there is data missing the system will either generate an error or assign default values. 

### Item ID, ID, Key, Ticket, Story ID
### Title, Summary, Name, Story
### Description, Details, Body, User Story
### Business Area, Domain, Department, BA
### Business Area Priority, BA Priority
### Product Team, Team, Squad, Pod
### Product Team Priority, Team Priority
### Effort, Story Points, SP, Size (numeric or XS/S/M/L/XL)


## Application Link

https://finalproject-jennahughes-diww6fzynhivpza7h5fyqw.streamlit.app/
