## The purpose of this project is as follows:
This strips and sorts a query that has employment data spanning campuses at CUNY.
## Here's some back story on why I needed to build this:
It is sometimes important, for reasons relating to fair administration of rights under the collective bargaining agreement, to have a history of employment at CUNY by campus by title and in a way that's more accessible than using Job Summary. This satisfies that criteria.
## This project leverages the following libraries:
matplotlib, pandas, pywin32, tabulate, xlwings
## In order to use this, you'll first need do the following:
The user must have access to CUNY and must construct a job history query using the EE_JOB record. I won't go into specifics but I'll share in HCMQ if there is anyone curious about the construction of the query I'm using.
## The expected frequency for running this code is as follows:
As Needed