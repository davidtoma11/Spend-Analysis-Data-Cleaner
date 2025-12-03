# Supplier Entity Resolution & Spend Analysis POC 

### Project Objective
The Procurement department faced a challenge with a fragmented database ("dirty data"), containing multiple inconsistent entries for the same suppliers. This Proof of Concept (POC) automates the data cleaning and deduplication process (**Entity Resolution**) to enable accurate **Spend Analysis**.

**The Challenge:**
Raw input data from the client was matched against an external database, returning 4-5 potential candidates per row. The goal was to algorithmically identify the "Golden Record" (the correct entity), preserve it, and eliminate the noise.

---

### The Logic: "Smart Scoring" Engine

Instead of relying on simple text matching, this script mimics the deductive process of a human analyst using a tiered scoring system.

#### 1. Standardization & Geo-Filtering 
* **Data Cleaning:** Removes invisible characters and standardizes text formats.
* **The Golden Rule:** `Input Country` must equal `Candidate Country`.
    * *Why?* A legal entity is defined by its tax registration country. Even if names are identical, a company in Germany is not the same legal entity as one in Romania. Mismatches are automatically rejected.

#### 2. The Smart Score Calculation �
We calculate a composite similarity score (0-100) based on:
* **Base Score:** Fuzzy name matching using `SequenceMatcher`.
* **Location Bonuses:** To differentiate between branches, we award extra points:
    * `+5 points` if Country matches.
    * `+3 points` if Region matches.
    * `+1 point` if City matches.

#### 3. Risk & Gap Analysis (The Decision Matrix) 
To determine the final winner, the algorithm analyzes the "Gap" between the top candidate and the runner-up:

* 🟢 **GREEN (Clear Winner):** A valid candidate exists with a score gap > 10 points compared to the 2nd place. **Auto-selected.**
* 🟡 **YELLOW (Close Call):** Multiple valid candidates exist with very close scores (gap <= 10 points). **All are retained for manual review.**
* 🔴 **RED (Forced Match):** No valid candidates found. The "best of the worst" is selected to preserve the spend record, but marked as high risk.


#### 4. Human verify (10-minutes fix)
While the algorithm automated the heavy lifting, the **"Yellow"** candidates (close calls) required a final human decision. Instead of a time-consuming deep dive, I applied a rapid validation strategy.

* **The Strategy:** I performed a visual scan focusing on a single, high-precision attribute: the **Postcode**.
* **The Decision:** In cases where names were similar but scores were close, if the **Postcode** matched, the candidate was immediately validated as the correct entity.
* **Efficiency:** Because the list was already filtered and color-coded, this manual verification took **less than 10 minutes** for the entire dataset.

The dataset is now ~95% resolved. The Data Analyst only has a few final checks to perform:
1.  Review the **"Red"** (Forced Match) rows to decide if they should be kept or deleted.
2.  The **"Green"** and resolved **"Yellow"** rows are considered final and ready for Spend Analysis.
---

### Output Files

The script generates two distinct Excel reports:

**1. Audit View (`file_1`)**
* **Purpose:** Full transparency.
* **Content:** Shows all 5 candidates per company with detailed color coding (Green/LightGreen/Red). No rows are hidden.

**2. Executive List (`file_2`)**
* **Purpose:** Actionable, clean data (One row per company).
* **Logic:**
    * **Dark Green:** Clear winners.
    * **Yellow:** Close calls (tie-breakers).
    * **Red:** Forced matches.
    * *Removed:* Lower-scoring valid alternatives are dropped to clean the list.

---

###  Requirments
* **Pandas** (Data Manipulation & Filtering)
* **Difflib** (SequenceMatcher for Fuzzy Logic)
* **OpenPyXL** (Excel Styling & Color Coding)

###  How to Run

1.  **Clone the repository:**
    ```bash
    git clone [https://github.com/your-username/supplier-entity-resolution-poc.git](https://github.com/your-username/supplier-entity-resolution-poc.git)
    ```

2.  **Install dependencies:**
    ```bash
    pip install pandas openpyxl
    ```

3.  **Run the script:**
    ```bash
    python entity_resolution.py
    ```

---

### Business Impact
This automation reduces manual data cleaning time by approximately **90%**. The analyst can now focus solely on the **Yellow** (ambiguous) and **Red** (high risk) entries, while the **Green** entries are processed with high confidence.
