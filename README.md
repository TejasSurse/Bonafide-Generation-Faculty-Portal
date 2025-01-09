# 🎓 Bonafide Generation and Management System 🚀

## 🌟 Overview
The **Bonafide Generation and Management System** is a modern, efficient platform that streamlines the process of generating and managing bonafide certificates for students. 📝✨ 

## 🌐 Live Links
- **Faculty Portal**: [Faculty Live Website](https://bonafide-generation-faculty-porta.vercel.app/)  
- **Faculty GitHub Repository**: [Faculty Portal Repo](https://github.com/TejasSurse/Bonafide-Generation-Faculty-Portal)  

- **Students Portal**: [Students Live Website](https://bonafide-generation-platform-lpqd.vercel.app/)  
- **Students GitHub Repository**: [Students Portal Repo](https://github.com/TejasSurse/Bonafide-Generation-Platform)  
-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------



### Key Features:
- 🎯 **Student Access**:  
  - Students can log in using their PRN and a password (first four characters of their first name + birth year).  
  - Instantly generate, download, and print their bonafide certificates. 🖨️📥  

- 🎓 **Faculty Portal**:  
  - Download an Excel format template for managing student records. 📄⬇️  
  - Upload the completed Excel sheet to update or add student data. The system auto-detects PRNs:  
    - If the PRN exists, it updates the data. 🔄  
    - If the PRN is new, it creates a new record. 🆕  

---

## 🛠️ Tech Stack
- **Backend**: Node.js, Express.js  
- **Frontend**: EJS, HTML, CSS, JavaScript  
- **Database**: MySQL (hosted on TiDB Cloud)  
- **Testing**: [Hoppscotch](https://hoppscotch.io)  
- **Deployment**: [Vercel](https://vercel.com)  

---

## 🚀 Features Breakdown
### 🎓 Student Portal:
1. **Login**:  
   - Enter PRN and password (`<First 4 letters of First Name> + <Birth Year>`). 🔐  
2. **Bonafide Certificate**:  
   - Auto-generate personalized certificates with student details. 📝  
   - Options to download or print the certificate. 📥🖨️  

### 🏫 Faculty Portal:
1. **Download Template**:  
   - Click to download the Excel format for uploading student data. ⬇️📄  
2. **Upload Data**:  
   - Upload the prepared Excel file.  
   - Auto-detects and updates/creates records in the database. 🔄📊  

---

## ⚙️ Installation and Setup

1. **Clone the repository**:  
   ```bash
   git clone https://github.com/yourusername/bonafide-management-system.git
   cd bonafide-management-system
Install dependencies:

bash
Copy code
npm install
Configure database:

Set up a MySQL database on TiDB Cloud.
Update the database connection details in .env.
plaintext
Copy code
DB_HOST=<your_database_host>
DB_USER=<your_database_user>
DB_PASSWORD=<your_database_password>
DB_NAME=<your_database_name>
Run the server:

bash
Copy code
npm start
Access the app at http://localhost:3000.

🧪 Testing
Use Hoppscotch to test the API endpoints.
Ensure endpoints like /upload-excel and /generate-bonafide are functioning as expected.
🌐 Deployment
Frontend & Backend: Deployed on Vercel.
Database: Hosted on TiDB Cloud.
📸 Screenshots
🖥️ Feature	💡 Description
Student Login Page	Enter PRN & password to log in.
Certificate Preview	Auto-generated bonafide certificate.
Faculty Portal	Download/upload Excel for student records.
💻 Contribution
Contributions are welcome! 🎉 Please follow these steps:

Fork the repository. 🍴
Create a new branch:
bash
Copy code
git checkout -b feature-name
Make your changes and commit:
bash
Copy code
git commit -m "Add feature-name"
Push to the branch:
bash
Copy code
git push origin feature-name
Create a pull request. 🛠️
📧 Contact
For any queries or suggestions, feel free to reach out! 📩

Developer: Tejas Sunil Surse
🌟 If you found this project helpful, please give it a ⭐ on GitHub!
