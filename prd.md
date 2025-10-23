# Product Requirements Document: Automated Debate Pairing Tool

## 1. Overview

An automated debate pairing tool built with Google Apps Script. It operates
within a Google Sheet, which serves as both the database and UI. All code is
self-contained in a single `Code.gs` file.

A one-time setup function initializes the sheet with necessary tabs and sample
data. The tool supports two debate formats: Team Policy (TP) and Lincoln-Douglas
(LD), managing rosters for debaters, judges, and rooms in dedicated tabs. It
prevents judging conflicts by tracking parent-child relationships.

The primary workflow is triggered weekly by an Admin using custom menu items.
Generated weekly sheets serve as the official record. The target scale is ~50
debaters, ~40 judges, ~20 rooms, and 1-3 admins with edit access.

## 2. The Problem

Manually creating fair and engaging debate pairings is a significant logistical
challenge for clubs. The process is time-consuming, prone to bias, and struggles
with variables like room availability, judge conflicts (e.g., parents judging
children), and skill levels. This leads to unbalanced debates, repetitive
matchups, and high administrative overhead.

## 3. User Personas

*   **Club Organizer (Admin):**
    *   Goal: To quickly generate fair, varied, and logistically sound debate
        pairings for each meeting using a familiar tool (Google Sheets).
*   **Club Member (Debater):**
    *   Goal: To have a fair and interesting debate against a similarly-skilled
        opponent in their chosen format.
*   **Judge (Parent):**
    *   Goal: To effectively judge a debate without conflicts of interest.

## 4. User Stories

### 4.1. Setup & Roster Management (Admin)

*   As an Admin, I want to use the Google Sheet as the single source of truth
    for all club data.
    *   As an Admin setting up the sheet for the first time, I want a one-click
        function to initialize the sheet, so that I can get started quickly.
    *   As an Admin using an already set-up sheet, I want to be protected from
        accidentally re-initializing it, to prevent data loss.
    *   As an Admin, I want the RSVP participant list to update automatically
        when rosters change, eliminating manual copy-pasting.
    *   As an Admin, I want to manage all debater, judge, and room information
        in a simple, centralized roster format.

### 4.2. Weekly Workflow (Admin)

*   As an Admin, I want to use simple menu commands to run the entire weekly
    meeting workflow.
    *   As an Admin, I want to see RSVPs in the Availability tab and be able to
        override them by editing cells.
    *   As an Admin, I want to generate all TP matches with a single click.
    *   As an Admin, I want to generate all LD matches with a single click.
    *   As an Admin, I want to clear all RSVPs after a meeting to prepare for
        the next one.
    *   As an Admin, I need the system to always generate a complete set of
        pairings, even if matchups aren't optimal. A suboptimal pairing is
        better than no pairing.
    *   As an Admin, I want to be able to manually adjust generated pairings in
        the sheet and immediately see if my changes created conflicts.

### 4.3. Participation & Viewing (Debaters & Judges)

*   As a Debater or Judge, I want to easily manage my participation and view my
    assignments.
    *   As a Debater or Judge, I need one place to set my attendance for the
        next meeting.
    *   As a Debater or Judge, I want a clear place to find my weekly
        assignment.

## 5. System Design & Logic

### 5.1. Sheet & Tab Structure

*   The tool operates in a single Google Sheet with a specific tab order.
*   **Permanent Tabs:** The first five tabs are permanent, in this order:
    1.  Availability
    2.  Match Summary
    3.  Debaters
    4.  Judges
    5.  Rooms
*   **Hidden Infrastructure Tabs:**
    *   `AGGREGATE_HISTORY`: A hidden tab used to persist all match data for
        historical analysis and soft constraint optimization.
*   **Generated Match Tabs:**
    *   Named by type and date (e.g., `TP 2025-07-29`).
    *   Sorted after the permanent tabs, with the most recent date first.
    *   For the same date, LD tabs appear before TP tabs.

### 5.2. Data Schemas

*   **Debaters Tab:** Name (Text, Unique), Debate Type ("TP" or "LD"), Partner
    (Text), Hard Mode ("Yes" or "No").
*   **Judges Tab:** Name (Text, Unique), Children's Names (Text,
    comma-delimited), Debate Type ("TP" or "LD").
*   **Rooms Tab:** Room Name (Text, Unique), Debate Type ("TP" or "LD").
*   **Availability Tab:**
    *   `Participant`: Populated by a formula combining unique names from
        Debaters and Judges.
    *   `Attending?`: Data validation dropdown ("Yes", "No", "Not responded").
        Defaults to "Not responded" via script when cleared.
*   **Match Summary Tab:**
    *   Must use formulas to aggregate data from the hidden history tab.
    *   Required Metrics per Debater: Total Matches, Aff Matches, Neg Matches,
        BYEs, Ironman Matches, Top 3 Judges (frequency), Top 3 Opponents
        (frequency).
*   **AGGREGATE_HISTORY Tab (Hidden):**
    *   Must capture sufficient detail to reconstruct any match: Date, Debate
        Type, Round, Debater Name, Role (Aff/Neg/BYE), Opponent Name, Judge
        Name(s), Room.
    *   *Note:* To support panel judging stats accurately, this table may need
        to denormalize data (e.g., one row per judge per debater).
*   **Generated Match Tabs:**
    *   TP Schema: `Aff Team | Neg Team | Judge(s) | Room`
    *   LD Schema: `Round | Aff Debater | Neg Debater | Judge(s) | Room`
    *   *Note:* Multiple judges (panels) must be listed in the single "Judge(s)"
        column, comma-delimited.

### 5.3. Menu Functions & Workflow Logic

*   An `onOpen` trigger creates a Club Admin menu.
*   **"Initialize Sheet"**: Creates permanent tabs with headers, sample data,
    formulas, and formatting. Must abort if it detects existing permanent tabs.
*   **"Generate TP Matches" / "Generate LD Matches"**:
    1.  Run data integrity checks (see 5.5). Stop on error.
    2.  Filter Availability tab for "Yes" to get active pool.
    3.  Run pairing logic (see 5.4).
    4.  Stop if a sheet for the current day/type already exists.
    5.  Write pairings to a new sheet.
    6.  Append new match data to `AGGREGATE_HISTORY`.
    7.  Re-sort all tabs.
*   **"Clear RSVPs"**: Resets all "Attending?" statuses to "Not responded".

### 5.4. Pairing Logic Constraints

*   **Hard Constraints (Required):**
    1.  Alert and stop if judges or rooms are insufficient for the minimum
        required matches.
    2.  Each match requires at least one unique, available judge.
    3.  Judges cannot officiate debates involving their own children.
    4.  Rooms and Judges must match the debate type.
*   **Special Cases:**
    *   **BYEs:** Odd number of teams/debaters results in a BYE. Assign to the
        participant with the fewest historical BYEs.
    *   **Ironman (TP):** A debater without their partner competes alone as an
        "Ironman" team (suffix name with `(IRONMAN)`).
    *   **Surplus Judges (Panels):** Extra available judges must be assigned to
        matches as additional judges to maximize feedback. Hard constraints
        still apply to all judges on a panel.
*   **Soft Constraints (Priorities):**
    1.  Pair participants with the same "Hard Mode" status.
    2.  Minimize rematches (vs. historical opponents).
    3.  Minimize re-judging (debaters judged by the same judge).
    4.  Preferentially assign judges to rooms they've used before.
*   **Failure Condition:** Relax soft constraints as needed to ensure all
    available debaters are paired.

#### 5.4.1. Special Logic for Lincoln-Douglas (LD)

LD debates use a two-round system on the same day.

1.  **Round 1:** Generate pairings using full history.
2.  **Round 2:** Generate pairings using full history *plus* the results of
    Round 1 (in-memory).
    *   *Constraint:* A debater cannot receive a BYE in both Round 1 and Round 2
        of the same event.
    *   *Constraint:* Resources (Judges/Rooms) should be reshuffled between
        rounds to maximize variety, provided hard constraints are met.
3.  **Finalize:** Output both rounds to the same generated sheet.

### 5.5. Data Validation & Integrity

The tool must use both script-based pre-checks (alert and halt on failure) and
continuous conditional formatting (visual red flags) to ensure data integrity.

*   **Roster Integrity (Pre-check & Conditional Formatting):**
    *   Room names must be unique.
    *   A person cannot be listed as both Judge and Debater.
    *   Judge's "Children's Names" must exist in the Debaters roster.
    *   TP Partnerships must be reciprocal and have matching "Hard Mode" status.
*   **Match Sheet Integrity (Conditional Formatting on Generated Sheets):**
    *   Must highlight any Participant, Room, or Judge not found in their
        respective rosters.
    *   *Note:* Judge validation must handle comma-delimited lists (panels),
        ensuring *every* judge in the cell exists in the roster.
    *   **TP:** Highlight if a debater or judge is assigned to >1 match.
    *   **LD:** Highlight if a debater is assigned to >1 match *per round*, or a
        judge is assigned to >1 match *total* (across both rounds).

### 5.6. User Experience (UX) Standards

To ensure usability by non-technical admins: * All sheets must have frozen top
rows so headers remain visible. * All data tables must use alternating row
colors (banding) for readability. * Columns should auto-resize to fit content
where appropriate. * Critical validation errors must use a high-contrast color
(e.g., light red background).

### 5.7. Initial Sample Data

*   Debaters:
    *   LD: Abraham Lincoln, Stephen A. Douglas, Clarence Darrow, William
        Jennings Bryan, William F. Buckley Jr., Gore Vidal, Christopher
        Hitchens, Tony Blair, Jordan Peterson, Slavoj Žižek, Lloyd Bentsen, Dan
        Quayle, Richard Dawkins, Rowan Williams, Diogenes
    *   TP Teams: Noam Chomsky, Michel Foucault; Harlow Shapley, Heber Curtis;
        Muhammad Ali, George Foreman; Richard Nixon, Nikita Khrushchev; Thomas
        Henry Huxley, Samuel Wilberforce; John F. Kennedy, David Frost; Bob
        Dole, Bill Clinton
*   Judges:
    *   TP: Howard K. Smith (Child: John F. Kennedy), Fons Elders (Children:
        Noam Chomsky, Michel Foucault), John Stevens Henslow (Child: Samuel
        Wilberforce), Jim Lehrer (Child: Bill Clinton), Judy Woodruff, Tom
        Brokaw, Frank McGee, Quincy Howe
    *   LD: John T. Raulston (Child: Clarence Darrow), Rudyard Griffiths (Child:
        Jordan Peterson), Stephen J. Blackwood, Brit Hume, Jon Margolis, Bill
        Shadel, Judge Judy
*   Rooms:
    *   LD: Room 101, Room 102, Room 103, Room 201, Sanctuary right, Sanctuary
        left, Pantry
    *   TP: Chapel, Library, Music lounge, Cry room, Office, Office hallway
