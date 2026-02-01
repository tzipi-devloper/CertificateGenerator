# Certificate Generator System

An automated certificate generation system that creates personalized PDF certificates for employees based on CSV data processing.

## 📋 Project Requirements

This C# application performs the following operations:

1. **Data Loading**: Load employee data from CSV file
2. **Data Cleaning**: Remove duplicate rows and fix name formatting
3. **Calculation & Filtering**: Calculate final score (60% practical + 40% theoretical)
4. **Data Storage**: Save processed data in appropriate format for certificate generation
5. **Certificate Generation**: Generate PDF certificates for qualified users (final score ≥ 70)

## 📊 Scoring Logic

- **Final Score Calculation**: 60% Practical Score + 40% Theoretical Score
- **Pass Threshold**: ≥ 70 points
- **Leadership Qualification**: > 90 points

## 📄 Certificate Templates

### High Achievers (Score > 90)
```
Phone: [Phone Number]
Email: [Email Address]
To: [Full Name], [Department]

We are pleased to inform you that you have successfully completed the training.
Your final score is [Score].
You are qualified for a departmental technology leadership position.

Best regards,
Company Management
Hogery Control & Compliance
```

### Standard Pass (Score 70-90)
```
Phone: [Phone Number]
Email: [Email Address]
To: [Full Name], [Department]

We regret to inform you that you did not pass the training,
and unfortunately, no suitable position was found for you.

Best regards,
Company Management
Hogery Control & Compliance
```

## 🏗️ Architecture & Features

### Core Components
- **Employee.cs**: Employee data model with score calculation
- **PdfGenerator.cs**: PDF generation using Word COM Interop
- **Logger.cs**: Comprehensive logging system
- **Program.cs**: Main application logic with LINQ processing

### Key Features
- ✅ LINQ-based data processing
- ✅ Duplicate removal and name formatting
- ✅ Word Mail Merge Field integration
- ✅ Resilient error handling (continues processing even if individual employee fails)
- ✅ Comprehensive logging system
- ✅ Automatic output directory creation
- ✅ Clean filename generation

## 🚀 Getting Started

### Prerequisites
- Visual Studio 2019 or later
- .NET Framework 4.7.2
- Microsoft Word (for COM Interop)
- Microsoft.Office.Interop.Word NuGet package

### Installation
1. Clone the repository
2. Open `CertificateGenerator.sln` in Visual Studio
3. Restore NuGet packages
4. Ensure `Data.csv` and `Template.docx` are in the project directory
5. Set "Copy to Output Directory" to "Copy always" for both files
6. Build and run the project

### Required Files
- `Data.csv` - Employee data (FirstName, LastName, Department, TheoryScore, PracticalScore)
- `Template.docx` - Word template with MERGEFIELD placeholders

### MERGEFIELD Placeholders
- `FullName` - Employee's full name
- `Department` - Employee's department
- `Phone` - Contact phone number
- `Email` - Employee's email address
- `BodyText` - Dynamic message based on score

## 📁 Output

- **PDF Files**: Generated in `Output/` directory with format `FirstName_LastName.pdf`
- **Logs**: Detailed process logs in `application.log`

## 🛡️ Production-Ready Improvements

For production deployment, the following enhancements were implemented:

### 1. Error Resilience
- Individual employee processing wrapped in try-catch
- System continues processing even if one employee fails
- Detailed error logging for troubleshooting

### 2. Data Validation
- CSV format validation
- Required file existence checks
- Score range validation
- Name format standardization

### 3. Logging & Monitoring
- Timestamped log entries
- Success/failure counters
- Process completion summary
- File operation logging

### 4. Resource Management
- Proper COM object disposal
- Memory cleanup for Word application
- Exception handling in finally blocks

### 5. File Safety
- Clean filename generation (removes invalid characters)
- Output directory auto-creation
- Template file read-only access

## 🔧 Technical Implementation

### Data Processing Pipeline
```csharp
var qualifiedEmployees = File.ReadAllLines(csvPath)
    .Skip(1) // Skip header
    .Select(line => line.Split(','))
    .Where(cols => cols.Length >= 5 && !string.IsNullOrWhiteSpace(cols[0]))
    .Select(cols => new Employee(cols))
    .GroupBy(emp => emp.FullName)
    .Select(g => g.First()) // Remove duplicates
    .Where(emp => emp.FinalScore >= 70)
    .ToList();
```

### Score Calculation
```csharp
FinalScore = (practical * 0.6) + (theory * 0.4);
```

### PDF Generation
- Uses Microsoft Word COM Interop
- Mail merge field replacement
- Automatic PDF export
- Document cleanup and disposal

## 📊 Sample Data Format

```csv
FirstName,LastName,Department,TheoryScore,PracticalScore
John,Doe,IT,85,92
Jane,Smith,HR,78,88
```

## 🤖 AI Assistance

This project was developed with AI assistance for:
- Code optimization and best practices
- Error handling implementation
- Documentation generation
- LINQ query optimization

## 📝 License

This project is developed for Hogery Control & Compliance.

## 🔗 Repository

Full source code available at: https://github.com/tzipi-devloper/CertificateGenerator