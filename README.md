# debatematcher

This repository contains the necessary files for an automated debate pairing
tool designed to run within a Google Sheet. The tool is built with Google Apps
Script and helps club organizers manage rosters, track availability, and
generate fair and balanced debate pairings for Team Policy (TP) and
Lincoln-Douglas (LD) formats.

## Files

### `prd.md` - The Blueprint

This file is a comprehensive **Product Requirements Document (PRD)**. It was
crafted to serve a dual purpose:

1.  **As a standard PRD:** It lays out the vision, user stories, functional
    requirements, and technical specifications for the project.
2.  **As a detailed prompt for a Software Engineering AI agent:** The level of
    detail, including specific data schemas, pairing logic, validation rules,
    and even sample data, was designed to be fed to an AI agent to generate the
    corresponding `code.gs` script. It serves as an example of how to
    effectively prompt an agent to produce a complete and functional piece of
    software.

### `code.gs` - The Script

This is the **Google Apps Script** file that powers the entire tool. It is a
self-contained script that should be added to a Google Sheet to enable the
debate pairing functionality. It handles everything from creating the user
interface (a custom menu) to managing the data and executing the complex pairing
logic.

## How to Use

To use the Automated Debate Matcher, follow these steps:

1.  **Create a new Google Sheet.** This will be your main interface for managing
    the debate club.
2.  Open the Apps Script editor by navigating to **Extensions > Apps Script**.
3.  Copy the entire content of the `code.gs` file from this repository.
4.  Paste the copied code into the script editor in your Google Sheet,
    completely replacing any default code (like an empty `myFunction`).
5.  Click the **Save project** icon (looks like a floppy disk).
6.  Return to your Google Sheet and **reload the page**.
7.  A new custom menu named **"Club Admin"** should now appear in the menu bar.
8.  Click **Club Admin > Initialize Sheet**. This will automatically set up all
    the necessary tabs (`Availability`, `Debaters`, `Judges`, `Rooms`) with the
    correct headers and sample data.

Your sheet is now ready to use! You can customize the rosters in the `Debaters`,
`Judges`, and `Rooms` tabs. When you're ready to create pairings, update the
`Availability` tab and use the "Generate Matches" functions from the "Club
Admin" menu.
