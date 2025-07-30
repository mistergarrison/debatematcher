# **Product Requirements Document: Automated Debate Pairing Tool**

## **1\. Overview**

An automated debate pairing tool built with Google Apps Script. It operates
within a Google Sheet, which serves as both the database and UI. All code is
self-contained in a single `Code.gs` file.

A one-time setup function initializes the sheet with necessary tabs and sample
data. The tool supports two debate formats: Team Policy (TP) and Lincoln-Douglas
(LD), managing rosters for debaters, judges, and rooms in dedicated tabs. It
prevents judging conflicts by tracking parent-child relationships.

The primary workflow is triggered weekly by an Admin using custom menu items.
Generated weekly sheets serve as the official record and match history. The
target scale is ~50 debaters, ~40 judges, ~20 rooms, and 1-3 admins with edit
access.

## **2\. The Problem**

Manually creating fair and engaging debate pairings is a significant logistical
challenge for clubs. The process is time-consuming, prone to bias, and struggles
with variables like room availability, judge conflicts (e.g., parents judging
children), and skill levels. This leads to unbalanced debates, repetitive
matchups, and high administrative overhead.

## **3\. User Personas**

*   **Club Organizer (Admin):**
    *   **Goal:** To quickly generate fair, varied, and logistically sound
        debate pairings for each meeting using a familiar tool (Google Sheets).
*   **Club Member (Debater):**
    *   **Goal:** To have a fair and interesting debate against a
        similarly-skilled opponent in their chosen format.
*   **Judge (Parent):**
    *   **Goal:** To effectively judge a debate without conflicts of interest.

## **4\. User Stories**

### **4.1. Setup & Roster Management (Admin)**

*   **\[EPIC\] As an Admin, I want to use the Google Sheet as the single source
    of truth for all club data.**
    *   **\[User Story\] As an Admin setting up the sheet for the first time,**
        I want a one-click function to initialize the sheet, so that I can get
        started quickly.
    *   **\[User Story\] As an Admin using an already set-up sheet,** I want to
        be protected from accidentally re-initializing it, to prevent data loss.
    *   **\[User Story\] As an Admin,** I want the RSVP participant list to
        update automatically when rosters change, eliminating manual
        copy-pasting.
    *   **\[User Story\] As an Admin,** I want to manage all debater, judge, and
        room information in a simple, centralized roster format.

### **4.2. Weekly Workflow (Admin)**

*   **\[EPIC\] As an Admin, I want to use simple menu commands to run the entire
    weekly meeting workflow.**
    *   **\[User Story\] As an Admin,** I want to see RSVPs in the Availability
        tab and be able to override them by editing cells.
    *   **\[User Story\] As an Admin,** I want to generate all TP matches with a
        single click.
    *   **\[User Story\] As an Admin,** I want to generate all LD matches with a
        single click.
    *   **\[User Story\] As an Admin,** after a meeting, I want a single action
        to clear all RSVPs to prepare for the next one.
    *   **\[User Story\] As an Admin,** I need the system to always generate a
        complete set of pairings, even if matchups aren't optimal. A suboptimal
        pairing is better than no pairing.
    *   **\[User Story\] As an Admin,** I want to be able to manually adjust
        generated pairings in the sheet.

### **4.3. Participation & Viewing (Debaters & Judges)**

*   **\[EPIC\] As a Debater or Judge, I want to easily manage my participation
    and view my assignments.**
    *   **\[User Story\] As a Debater or Judge,** I need one place to set my
        attendance for the next meeting.
    *   **\[User Story\] As a Debater or Judge,** I want a clear place to find
        my weekly assignment.

## **5\. System Design & Logic**

### **5.1. Sheet & Tab Structure**

*   The tool operates in a single Google Sheet with a specific tab order.
*   **Permanent Tabs:** The first four tabs are permanent, in this order:
    1.  Availability
    2.  Debaters
    3.  Judges
    4.  Rooms
*   All sheets must have a frozen and bolded header row.
*   **Generated Match Tabs:**
    *   Named by type and date (e.g., `TP 2025-07-29`), they contain the final
        pairings and together form the match history.
    *   They are sorted after the permanent tabs, with the most recent date
        first.
    *   For the same date, LD tabs appear before TP tabs.

### **5.2. Data Schemas & Formatting**

*   **Debaters Tab:**
    *   Columns: Name (Text), Debate Type ("TP" or "LD"), Partner (Text), Hard
        Mode ("Yes" or "No").
*   **Judges Tab:**
    *   Columns: Name (Text), Children's Names (Text, comma-delimited), Debate
        Type ("TP" or "LD").
*   **Rooms Tab:**
    *   Columns: Room Name (Text), Debate Type ("TP" or "LD").
*   **Availability Tab:**
    *   The **Participant** column is populated by a sheet formula that combines
        and sorts unique names from the **Debaters** and **Judges** tabs.
    *   The **Attending?** column must have a data validation dropdown with
        "Yes", "No", and "Not responded" (blanks allowed).
    *   A conditional format must highlight the **Attending?** cell in light red
        if it's blank but the adjacent **Participant** cell is not.
    *   The **Attending?** column's default value is "Not responded" if a
        participant is present, otherwise it's blank. This logic is applied by
        both `initializeSheet` and `clearRsvps`.

### **5.3. Menu Functions & Workflow Logic**

*   An `onOpen` trigger creates a **Club Admin** menu.
*   **"Initialize Sheet":** Creates the four permanent tabs with headers and
    sample data (see section 5.6). It also sets up all required formulas and
    formatting. It won't overwrite existing tabs.
*   **"Generate TP Matches" / "Generate LD Matches":** Each item will:
    1.  Run data integrity checks (see 5.5). The process stops on any error.
    2.  Filter for "Yes" in the **Availability** tab to get available
        participants and resources.
    3.  Run the pairing logic.
    4.  Show an error and stop if a sheet for the current day and type already
        exists, to prevent overwrites.
    5.  Write pairings to a new sheet named by type and date (e.g., `TP
        YYYY-MM-DD`).
    6.  Re-sort all tabs automatically.
*   **"Clear RSVPs":** Sets the **Attending?** value to "Not responded" for
    every participant.

### **5.4. Pairing Logic Constraints**

*   **Hard Constraints (Required):**
    1.  Stop with an alert if judges or rooms are insufficient for the number of
        matches.
    2.  Each match requires one unique, available judge.
    3.  Judges cannot officiate debates involving their own children.
    4.  Rooms must match the debate type.
    5.  Judges must match the debate type.
*   **Special Cases:**
    *   **Odd Number of Debaters/Teams:** An odd number of participants results
        in a BYE. The BYE is assigned to the participant/team with the fewest
        historical BYEs to ensure fairness. BYEs don't need a judge or room.
    *   **Incomplete TP Teams:** A TP debater whose partner is absent will
        compete as an "Ironman" team (e.g., "Debater Name (IRONMAN)") and is
        paired normally.
*   **Soft Constraints (Priorities):**
    1.  Pair participants with the same "Hard Mode" status.
    2.  Minimize rematches (vs. historical opponents).
    3.  Minimize re-judging (debaters judged by the same judge).
    4.  Preferentially assign judges to rooms they've used before.
*   **Failure Condition:** Relax soft constraints as needed to ensure all
    available debaters are paired.

### **5.4.1. Special Logic for Lincoln-Douglas Pairings**

LD debates use a two-round system to ensure varied matchups and fair BYE
distribution. The logic is sequential:

1.  **Generate Round 1:**

    *   Generate Round 1 pairings using the full match history. This includes
        assigning a BYE if needed.
    *   Assign judges and rooms from the full available pool.

2.  **Update History In-Memory:**

    *   Update a *copy* of the match history in-memory with Round 1 results
        (opponents, judges, BYEs). This is critical for fair Round 2 pairing.

3.  **Generate Round 2:**

    *   Generate a new set of Round 2 pairings using the updated in-memory
        history. This prevents repeat BYEs and matchups.
    *   Assign judges and rooms from the same full pool. Resources can be
        reused, but soft constraints still apply to encourage variety.

4.  **Combine and Finalize:**

    *   The final output sheet will list assignments for both rounds.

### **5.5. Data Validation**

The "Generate Matches" function must run data validation. On failure, it must
alert the user with the specific error and location, then stop.

For better UX, these rules will also be implemented as conditional formatting
(highlighting invalid rows in light red). The script must re-apply these rules
on every open to ensure they are always active.

**Implementation Note:** When applying these rules via App Script, be aware that
the `sheet.setConditionalFormatRules()` method overwrites all existing rules on
that sheet. The correct approach is to build a single array containing all
desired `ConditionalFormatRule` objects for a sheet and then make only one call
to `setConditionalFormatRules()` with that complete array. Custom formulas for
conditional formatting cannot directly reference other sheets; therefore, any
validation that requires cross-sheet lookups (e.g., checking if a judge exists
in the "Judges" roster) must use the `INDIRECT()` function.

**Validation Rules:**

*   **On the Rooms tab:**
    *   Room names must be unique.
*   **On the Judges tab:**
    *   A person cannot be both a Judge and a Debater.
    *   All names in "Children's Names" must exist in the Debaters roster.
*   **On the Debaters tab:**
    *   **Partnership Consistency (for all TP debaters):**
        *   Partner must exist in the Debaters roster.
        *   Partnerships must be reciprocal.
        *   Partners must have the same "Hard Mode" setting.
*   **On Generated Match Tabs (e.g., "TP 2025-07-29"):**
    *   **Sheet Schemas & Sort Order:**
        *   **TP Sheets:** `Aff | Neg | Judge | Room`
        *   **LD Sheets:** `Round | Aff | Neg | Judge | Room`
            *   Data rows are sorted by `Round` (asc), then `Room` (asc).
    *   **Referential Integrity & Uniqueness:** Apply conditional formatting to
        new match sheets to verify data integrity:
        *   All participants and resources must exist in their respective
            rosters. Highlight missing entries.
        *   **TP:** Each debater and judge is assigned to only one match.
            Highlight duplicate assignments.
        *   **LD:** Each debater is in two matches (one per round). Each judge
            is assigned to one matchup (covering both rounds). Highlight judges
            assigned to more than one matchup.
        *   A pre-check must halt with an error if any debater is assigned to
            more than one match per round (a critical logic failure).
    *   **Special Case Handling:** Validation must not flag these valid cases:
        *   `"BYE"` as an opponent is valid.
        *   `"(IRONMAN)"` names are valid; validation should check the base name
            against the roster.

### **5.6. Initial Sample Data**

*   **Debaters:**
    *   **LD:** Abraham Lincoln, Stephen A. Douglas, Clarence Darrow, William
        Jennings Bryan, William F. Buckley Jr., Gore Vidal, Christopher
        Hitchens, Tony Blair, Jordan Peterson, Slavoj Žižek, Lloyd Bentsen, Dan
        Quayle, Richard Dawkins, Rowan Williams, Diogenes
    *   **TP Teams:** Noam Chomsky, Michel Foucault; Harlow Shapley, Heber
        Curtis; Muhammad Ali, George Foreman; Richard Nixon, Nikita Khrushchev;
        Thomas Henry Huxley, Samuel Wilberforce; John F. Kennedy, David Frost;
        Bob Dole, Bill Clinton
*   **Judges:**
    *   **TP:** Howard K. Smith (Child: John F. Kennedy), Fons Elders (Children:
        Noam Chomsky, Michel Foucault), John Stevens Henslow (Child: Samuel
        Wilberforce), Jim Lehrer (Child: Bill Clinton), Judy Woodruff, Tom
        Brokaw, Frank McGee, Quincy Howe
    *   **LD:** John T. Raulston (Child: Clarence Darrow), Rudyard Griffiths
        (Child: Jordan Peterson), Stephen J. Blackwood, Brit Hume, Jon Margolis,
        Bill Shadel, Judge Judy
*   **Rooms:**
    *   **LD:** Room 101, Room 102, Room 103, Room 201, Sanctuary right,
        Sanctuary left, Pantry
    *   **TP:** Chapel, Library, Music lounge, Cry room, Office, Office hallway

### **5.7. User Notifications**

*   Use pop-up alerts (`ui.alert`) only for critical errors that halt a process.
*   Do not use alerts for success messages. The visual change in the sheet is
    sufficient feedback.

## **6\. Technical Specifications**

*   **Platform:** Google App Script.
*   **Code:**
    *   All logic is contained in a single `Code.gs` file.
    *   Code must be modular, with complex logic broken into small,
        single-purpose functions.
    *   Every function requires a detailed and helpful JSDoc comment.
    *   Use inline comments to clarify complex code.
*   **UI:** The Google Sheet.
*   **Data Storage:** The Google Sheet tabs.

## **7\. Success Metrics**

*   100% of meetings are paired by the script.
*   Pairing generation time is reduced from >30 minutes to <1 minute.
*   100% of parent-child judging conflicts are prevented.
*   Soft pairing rules are applied in >95% of matches.
*   Admins average <3 manual edits to generated pairings per meeting.
