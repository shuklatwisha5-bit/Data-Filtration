# Data-Filtration
Read, Filter &amp; Extract Clean Data from Excel using Apache POI
# Java Excel Data Processor (Maven Backend)

This is a robust, Maven-based Java project designed to perform essential data processing tasks—specifically **data cleaning (Imputation)** and **compound filtering**—on Microsoft Excel (.xlsx) files. The project utilizes the **Apache POI** library for programmatic interaction with Excel documents.

## 📊 Project Utilities

The backend contains two primary utilities, both located in the com.data.filter package, designed to be run independently via the command line.

### 1. ExcelImputer (Data Cleaning)

* **Purpose:** To clean and standardize the input data for reliable analysis.
* **Logic:** Reads the input Excel file and replaces any blank, null, or non-numeric cells in score columns (`Math`, `Science`, `English`) with a default numeric value of **0.0**.

### 2. ExcelFilter (Compound Filtering)

* **Purpose:** To select specific data rows that meet complex, multi-criteria performance standards.
* **Logic:** Applies a logical **AND** condition across two score fields. For example, it selects a student only if **(Math Score >= 85 AND Science Score >= 85)**. The specific thresholds are defined within the source code.

## 🛠️ Development Environment & Tools

This section outlines the complete technology stack required to build and run the project.

| Category | Component/Tool | Detail |
| :--- | :--- | :--- |
| **Language** | **Java** | Primary programming language (JDK 17+ recommended). |
| **Build Tool** | **Apache Maven** | Manages dependencies (POI) and handles the build and execution life cycle. |
| **Core Library** | **Apache POI** | Provides the API for reading and writing modern Excel (.xlsx) files. |
| **IDE** | **IntelliJ IDEA / VS Code** | Recommended tools for development and debugging. |
| **Execution App** | **CLI** | Command Line Interface (Terminal/PowerShell) used to run Maven commands. |
| **Version Control** | **Git** | Used for tracking changes and managing source code versions on GitHub. |

## ⚙️ Setup and Execution

### Prerequisites

You must have the following software installed on your system:

1.  **Java Development Kit (JDK)**: Version 17 or compatible.
2.  **Apache Maven**.

### Project Structure
excel-processor-backend/ ├── src/ │ └── main/ │ └── java/ │ └── com/ │ └── data/ │ └── filter/ │ ├── ExcelImputer.java │ └── ExcelFilter.java ├── pom.xml └── README.md
### Execution Steps

1.  **Place Input File:**
    Ensure your input data file, named **`student_marks_input.xlsx`**, is placed directly in the root directory of this project.

2.  **Run from Command Line:**
    Navigate to the root directory in your terminal and use the appropriate Maven command:

    #### A. Run Data Imputer (Data Cleaning)

    This executes `ExcelImputer` and generates the clean file, `data_imputed.xlsx`.

    ```bash
    mvn clean compile exec:java -Dexec.mainClass="com.data.filter.ExcelImputer"
    ```

    #### B. Run Compound Filter (Data Selection)

    This executes `ExcelFilter` and generates the filtered results file, `dual_high_scorers_output.xlsx`.

    ```bash
    mvn clean compile exec:java -Dexec.mainClass="com.data.filter.ExcelFilter"
    ```
