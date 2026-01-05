Student Data Excel Processor & Dashboard 📊

A robust Java-based data processing application that utilizes Apache POI to handle Excel workbooks, perform data imputation (handling missing values), and filter student records based on custom academic thresholds. It includes a modern, responsive web dashboard for visualizing results.

🚀 Features

Excel Imputation: Automatically detects missing grade cells and imputes them (defaulting to 0.0) to prevent calculation errors.

Advanced Filtering: Filter student records across multiple subjects (Math, Science, English) using configurable thresholds.

Apache POI Integration: Seamlessly reads and writes .xlsx files.

Maven Architecture: Standardized project structure for easy dependency management and builds.

Interactive Dashboard: A Tailwind CSS-powered GUI (index.html) to visualize backend processing logic and filter results dynamically.

📂 Project Structure

excel-imputer/
├── src/
│   ├── main/
│   │   ├── java/
│   │   │   └── com/data/filter/
│   │   │       ├── ExcelImputer.java   # Logic for handling missing data
│   │   │       └── ExcelFilter.java    # Logic for academic filtering
│   │   └── resources/
│   │       └── index.html              # Frontend visualization dashboard
├── student_marks_input.xlsx             # Source Excel data file
├── pom.xml                             # Maven configuration & dependencies
└── README.md                           # Project documentation


🛠️ Technical Stack

Backend: Java 11+

Build Tool: Maven 3.6+

Libraries: Apache POI (Excel Processing), Log4j

Frontend: HTML5, Tailwind CSS (CDN), JavaScript (ES6+)

⚙️ Setup & Installation

Prerequisites

Java Development Kit (JDK) 11 or higher.

Apache Maven installed.

An IDE (IntelliJ IDEA recommended).

1. Clone the Repository

git clone [https://github.com/yourusername/excel-imputer.git](https://github.com/yourusername/excel-imputer.git)
cd excel-imputer


2. Install Dependencies

Run the following command to download the required Apache POI libraries:

mvn clean install


3. Running the Backend

To execute the filter logic via terminal:

mvn exec:java -Dexec.mainClass="com.data.filter.ExcelFilter"


🖥️ Using the Dashboard

Navigate to src/main/resources/index.html.

Right-click the file in IntelliJ and select Open in Browser.

Features available in the GUI:

Threshold Adjustment: Change minimum Math/Science requirements.

Live Search: Filter students by name.

System Console: View simulated Maven build logs.

Imputation Markers: Note the 0.0* markers highlighting where the Java backend filled in missing data.

📝 Configuration (pom.xml)

The project relies on the following key dependencies in the pom.xml:

poi-ooxml: For reading/writing Excel files.

log4j-core: For systematic logging.

🤝 Contributing

Fork the Project.

Create your Feature Branch (git checkout -b feature/AmazingFeature).

Commit your Changes (git commit -m 'Add some AmazingFeature').

Push to the Branch (git push origin feature/AmazingFeature).

Open a Pull Request.

📄 License

Distributed under the MIT License. See LICENSE for more information.
