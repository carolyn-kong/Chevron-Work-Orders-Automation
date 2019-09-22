# Chevron-Work-Orders-Automation
A work order automation system that optimizes for worker productivity and automatically sends workers' schedules in real time.


## Inspiration
Chevron has many facilities that require different types of repairs. Scheduling specialized technicians for all of the complex process equipment between the many facilities in an optimized fashion can be difficult.

## What it does
Our project starts with a Google Chrome extension that links to a Google Form. In this Google Form, Chevron employees can submit work orders for equipment that needs to be maintained or repaired. These work orders are stored in an Excel sheet that is referenced by our Python algorithm. The algorithm prioritizes the work requests and assigns each task to a specialized worker, who is notified of their work schedule via SMS messaging.

## How we built it
Our team started by discussing the prioritization algorithm and user interface we wanted to implement. We then separated into two teams: Alex and Robert primarily worked on the prioritization algorithm and Carolyn and Audrey primarily worked on the user interface. 

Alex and Robert created a Python program that functions as follows: There are two workers, one that works in the morning and one that works in the evening, that have the same specialization and therefore the same list of tasks. The lists of tasks were optimized by how urgent they need to be fixed (their priority) as well as optimizing the travel time, time to complete a task, and how long the work order has been submitted.

Carolyn and Audrey created a Google Form to submit work orders in the same format as in the given data set. This Google Form is accessible via a Google Chrome Extension that was created with HTML, JSON, and JavaScript. The responses from the Google Form are automatically entered into a Google Sheet that is published to the web. Using this web address, Microsoft Excel is able to automatically update a spreadsheet that is referenced by the Python program.

## Challenges we ran into
Some challenges we ran into include:
- transferring work order submissions from the web to an Excel sheet automatically 
- developing an algorithm to prioritize work orders

## Accomplishments that we're proud of
- SMS messaging system

## What we learned
- basic Google Chrome Extensions
- Excel sync to web
- Python

## What's next for Chevron Work Orders Automation
- faster sync between Google Forms and Excel
- ability to mark work as completed
