# APDocumentGen

> **Automating patient document preparation for pediatric clinics through intelligent appointment parsing and dynamic PDF generation.**

APDocumentGen is a production Go CLI that streamlines the administrative workflow of a pediatric clinic by automatically generating the required patient paperwork before each appointment. Given a daily appointment schedule, the application analyzes each patient's visit reason, determines the necessary medical forms and educational materials, and produces a single merged PDF packet for every patient.

Designed for daily clinical use, APDocumentGen reduces repetitive manual preparation, helps ensure appointment-specific documentation is never missed, and allows staff to focus more on patient care rather than paperwork.

---

## ✨ Features

- 📅 Parse daily appointment schedules from CSV or Excel files
- 🩺 Analyze visit reasons using keyword-based medical rules
- 💉 Automatically determine required immunizations and screenings
- 📄 Generate a consolidated PDF packet for each patient
- 📚 Include educational handouts and informational documents based on visit type
- ⚡ Batch process an entire day's appointments or generate documents for an individual patient
- 🛠️ Easily extendable rule mappings for new visit types and medical requirements

---

## 🏥 Example Workflow

```text
Appointment Schedule
        │
        ▼
 Read CSV / Excel File
        │
        ▼
 Parse Patient Information
        │
        ▼
 Analyze Visit Reason
        │
        ▼
 Map Keywords → Required Documents
        │
        ▼
 Merge PDF Templates
        │
        ▼
 One Complete Patient Packet
```

For example, a visit labeled:

```text
2Y WCC
```

may automatically include:

- Immunization forms
- Developmental and behavioral screening questionnaires
- Vaccine Information Statements (VIS)
- Hearing and vision screening materials
- Age-specific educational handouts

All required documents are merged into a single PDF packet ready for the patient's appointment.

---

## 🛠️ Tech Stack

| Category | Technologies |
|----------|--------------|
| Language | Go |
| PDF Generation | `gofpdf` |
| CSV Parsing | `encoding/csv` |
| Excel Support | `excelize`, `xlsReader` |
| CLI | Go `flag` package |
| Data Processing | Go Standard Library |

---

## 📁 Project Structure

```text
APDocumentGen/
├── main.go
├── parser/
├── rules/
├── generator/
├── templates/      # Excluded from repository
├── output/
└── README.md
```

> **Note:** PDF templates used in production are intentionally excluded from this repository to protect clinic resources and comply with healthcare privacy and office security requirements.

---

## ⚙️ How It Works

The application follows a rule-based document generation pipeline.

1. Read a schedule exported from the clinic's appointment system.
2. Extract patient information including:
   - Name
   - Date of Birth
   - Appointment Time
   - Visit Reason
3. Parse the visit reason using keyword matching.
4. Map detected keywords to predefined medical requirements.
5. Collect all applicable document templates.
6. Merge templates into a single PDF packet.
7. Save the completed packet for printing or digital distribution.

Because document selection is driven by configurable keyword mappings, new visit types and clinic workflows can be added with minimal code changes.

---

## 🚀 Getting Started

### Prerequisites

- Go 1.20+
- PDF template directory (not included in this repository)

### Installation

```bash
git clone https://github.com/yourusername/APDocumentGen.git
cd APDocumentGen

go mod download
```

### Usage

Run the application against a daily appointment schedule:

```bash
go run main.go appointments.csv
```

or after building:

```bash
go build -o apdocumentgen

./apdocumentgen appointments.csv
```

The application processes each appointment and generates a consolidated PDF packet for every patient.

---

## 🔒 Privacy & Security

This project was developed for use in a pediatric healthcare environment.

To protect patient privacy and clinic resources:

- Production PDF templates are intentionally excluded from this repository.
- No patient data is stored within the application.
- Sample datasets should be anonymized before testing.
- This repository contains only the document generation logic—not protected clinical assets.

---

## 💡 Design Decisions

Rather than relying on hardcoded workflows for every appointment type, APDocumentGen uses keyword-based rules to dynamically determine the required documentation.

This approach makes the system:

- Easy to maintain
- Adaptable to changing clinical workflows
- Extensible for new visit types
- Significantly faster than manually assembling paperwork

---

## 🎯 Future Improvements

- 🌐 Develop a user-friendly web interface for non-technical staff
- ☁️ Deploy to the cloud for centralized access across multiple clinics
- 🤖 Implement an autonomous agent that monitors appointment exports and automatically generates patient packets
- ⚙️ Create a configurable rule editor for clinic administrators
- 🏥 Integrate with Electronic Health Record (EHR) systems
- 📧 Send notifications when document generation is complete

---

## 📈 Impact

APDocumentGen was built to solve a real operational challenge in a pediatric clinic by automating the repetitive task of preparing appointment-specific paperwork.

By transforming appointment schedules into ready-to-print patient packets, the tool reduces administrative overhead, minimizes missed documentation, and improves overall clinic efficiency while allowing staff to dedicate more time to patient care.

---

## 👨‍💻 Author

Developed by **Sanjay** as a production automation tool for pediatric clinical workflows.

If you found this project interesting, feel free to ⭐ the repository!
