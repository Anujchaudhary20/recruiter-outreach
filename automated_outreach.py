#!/usr/bin/env python3
"""
Automated Daily LinkedIn Recruiter Outreach System
Runs daily at 10:30 AM IST via GitHub Actions
Scrapes recruiters → Tailors resumes → Sends emails → Updates tracker
"""

import os
import json
import requests
from datetime import datetime, timedelta
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import openpyxl
from openpyxl.utils import get_column_letter

# Configuration
FIRECRAWL_API_KEY = os.getenv('FIRECRAWL_API_KEY', '')
GMAIL_EMAIL = os.getenv('ANUJ_EMAIL', 'Anuj04004@gmail.com')
GMAIL_APP_PASSWORD = os.getenv('GMAIL_APP_PASSWORD', '')
RESUME_TEMPLATE_PATH = os.getenv('RESUME_TEMPLATE_PATH', 'resume_template.txt')

# Your details
ANUJ_NAME = "Anuj Singh"
ANUJ_PHONE = os.getenv('ANUJ_PHONE', '+91-9528034629')
ANUJ_LINKEDIN = "linkedin.com/in/anujsingh-java"

# Test recruiters (will be replaced with real Firecrawl data)
TEST_RECRUITERS = [
    {
        'recruiter_name': 'Raj Kumar',
        'company': 'TechCorp India',
        'email': 'raj.kumar@techcorp.com',
        'location': 'Bengaluru',
        'job_description': 'Hiring Senior Java Developer with Spring Boot and Microservices experience'
    },
    {
        'recruiter_name': 'Priya Singh',
        'company': 'SoftDev Solutions',
        'email': 'priya@softdev.com',
        'location': 'Noida',
        'job_description': 'Looking for Java Backend Engineer with Kafka and distributed systems'
    },
    {
        'recruiter_name': 'Arun Patel',
        'company': 'CloudInc Technologies',
        'email': 'arun.patel@cloudinc.com',
        'location': 'Mumbai',
        'job_description': 'Recruiting Java Developer for AWS cloud projects'
    },
    {
        'recruiter_name': 'Neha Sharma',
        'company': 'DataSystems Ltd',
        'email': 'neha@datasystems.com',
        'location': 'Hyderabad',
        'job_description': 'Senior Java Engineer needed for real-time data processing'
    },
    {
        'recruiter_name': 'Vikram Singh',
        'company': 'InnovateTech',
        'email': 'vikram@innovatetech.com',
        'location': 'Pune',
        'job_description': 'Java Full Stack Developer for microservices architecture'
    },
]

def send_email_with_resume(email_data, resume):
    """Send email via Gmail SMTP"""
    try:
        msg = MIMEMultipart()
        msg['From'] = GMAIL_EMAIL
        msg['To'] = email_data['to']
        msg['Subject'] = email_data['subject']
        msg.attach(MIMEText(email_data['body'], 'plain'))
        
        # Connect and send via Gmail SMTP
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(GMAIL_EMAIL, GMAIL_APP_PASSWORD)
            server.send_message(msg)
        
        print(f"✅ Email sent to {email_data['to']}")
        return True
    except Exception as e:
        print(f"❌ Failed to send to {email_data['to']}: {e}")
        return False

def get_standard_resume():
    """Load standard resume template"""
    try:
        with open(RESUME_TEMPLATE_PATH, 'r') as f:
            return f.read()
    except:
        # Fallback resume if file not found
        return """ANUJ KUMAR SINGH
+91 9528034629 | Anuj04004@gmail.com | Bengaluru, India

SUMMARY
Java Backend Engineer with 4 years of experience in Spring Boot, Microservices, 
REST APIs, and Distributed Systems. Proficient in event-driven architecture, 
cloud-native deployments, and CI/CD.

SKILLS
Languages & Frameworks: Java 17, Java 21, Spring Boot, Spring Security, Spring Data JPA
Databases & Messaging: PostgreSQL, MySQL, Redis, Apache Kafka, Kafka Streams
Cloud & DevOps: AWS (EC2, Lambda, S3), Azure, Docker, Kubernetes, CI/CD pipelines
Testing & Tools: JUnit, Mockito, IntelliJ IDEA, GitHub, Jenkins, Prometheus, Grafana

WORK EXPERIENCE
Nagarro, Remote | July 2025 – Present
Senior Software Engineer
• Designed event-driven microservice modules using Spring Boot and Java 17
• Led Gradle-to-Maven migration on 100K+ LOC distributed system
• Embedded GitHub Copilot across SDLC, reducing implementation time by 45%
• Mentored 2 junior engineers on microservice design patterns

Capgemini, Bengaluru | Sep 2022 – July 2025
Software Engineer
• Owned end-to-end feature delivery, achieving zero critical incidents
• Architected 15+ RESTful APIs using Spring Boot and Swagger/OpenAPI
• Implemented Redis caching layer, reducing response time by 10%
• Refactored legacy Spring MVC modules, raising test coverage to 80%

EDUCATION
GLA University, Mathura
Bachelor of Computer Applications (BCA) | August 2019 – May 2022

CERTIFICATIONS
AWS Certified Cloud Practitioner
Microsoft Azure Fundamentals
Agile Software Development"""

def tailor_resume(recruiter_data):
    """Tailor resume to match job description"""
    jd = recruiter_data.get('job_description', '')
    company = recruiter_data.get('company', 'Company')
    
    resume = get_standard_resume()
    
    # Simple tailoring: highlight relevant keywords
    if 'kafka' in jd.lower():
        resume = resume.replace(
            'Kafka Streams',
            'Apache Kafka, Kafka Streams (HIGHLIGHTED)'
        )
    
    if 'microservice' in jd.lower():
        resume = resume.replace(
            'microservice modules',
            'microservice modules (RELEVANT TO THIS ROLE)'
        )
    
    if 'aws' in jd.lower():
        resume = resume.replace(
            'AWS (EC2, Lambda, S3)',
            'AWS (EC2, Lambda, S3, VPC) - EXPERT LEVEL'
        )
    
    return resume

def generate_email(recruiter_data, resume):
    """Generate personalized email"""
    name = recruiter_data.get('recruiter_name', 'Hiring Manager')
    company = recruiter_data.get('company', 'Company')
    jd = recruiter_data.get('job_description', '')
    
    # Extract detail from JD
    detail = jd.split('.')[0] if jd else 'for a Java position'
    
    email_body = f"""Hi {name},

I came across your recent hiring post at {company}, and I'm genuinely interested in the opportunity.

I'm a Senior Java Backend Engineer with 4 years of hands-on experience building high-scale systems using:
• Spring Boot & Microservices Architecture
• Apache Kafka, Redis, AWS/Azure
• System Design & Distributed Systems
• Docker, Kubernetes, CI/CD Pipelines

Your focus on {detail} resonates with my expertise, and I believe I can add significant value to your team.

I'm actively looking for a new challenge and available immediately to join. My attached resume highlights relevant projects and achievements.

Would love to discuss how I can contribute to your initiatives. Open to a quick call at your convenience.

Best regards,
Anuj Singh
{GMAIL_EMAIL}
{ANUJ_PHONE}
LinkedIn: {ANUJ_LINKEDIN}"""
    
    subject = f"Java Developer | 4 Years | Spring Boot | Immediate Joiner"
    
    return {
        'subject': subject,
        'body': email_body,
        'to': recruiter_data.get('email')
    }

def update_excel_tracker(recruiter_data):
    """Update Excel tracking sheet"""
    excel_tracker = "recruiter_tracking.xlsx"
    
    try:
        # Load or create workbook
        if os.path.exists(excel_tracker):
            wb = openpyxl.load_workbook(excel_tracker)
            ws = wb.active
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            
            # Add headers
            headers = [
                'Contact Name', 'Email', 'Company', 'Job Title', 'Location',
                'Date Sent', 'Status', 'Notes'
            ]
            for col, header in enumerate(headers, 1):
                ws[f'{get_column_letter(col)}1'] = header
        
        # Add new row
        row = ws.max_row + 1
        ws[f'A{row}'] = recruiter_data.get('recruiter_name')
        ws[f'B{row}'] = recruiter_data.get('email')
        ws[f'C{row}'] = recruiter_data.get('company')
        ws[f'D{row}'] = recruiter_data.get('location', '')
        ws[f'E{row}'] = datetime.now().strftime('%Y-%m-%d')
        ws[f'F{row}'] = 'Sent'
        ws[f'G{row}'] = 'Waiting for reply'
        
        wb.save(excel_tracker)
        print(f"✅ Excel tracker updated")
        
    except Exception as e:
        print(f"⚠️ Excel update error: {e}")

def scrape_linkedin_recruiters():
    """Scrape LinkedIn for India-based recruiters"""
    print("🔍 Searching for recruiters...")
    
    # Using test data for now
    # TODO: Replace with real Firecrawl integration
    recruiters = TEST_RECRUITERS
    
    print(f"✅ Found {len(recruiters)} recruiters")
    return recruiters

class OutreachAutomation:
    def __init__(self):
        self.recruiters = []
        self.sent_count = 0
        self.failed_count = 0
        
    def run(self):
        """Main execution"""
        print("\n" + "="*60)
        print("🚀 AUTOMATED RECRUITER OUTREACH - STARTED")
        print("="*60)
        print(f"⏰ Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        
        # Step 1: Scrape
        self.recruiters = scrape_linkedin_recruiters()
        
        if not self.recruiters:
            print("❌ No recruiters found")
            return
        
        # Step 2: Generate & Send emails
        print(f"\n📧 Generating and sending {len(self.recruiters)} emails...\n")
        
        for i, recruiter in enumerate(self.recruiters, 1):
            print(f"[{i}/{len(self.recruiters)}] Processing: {recruiter.get('recruiter_name')}")
            
            # Tailor resume
            resume = tailor_resume(recruiter)
            
            # Generate email
            email_data = generate_email(recruiter, resume)
            
            # Send
            if send_email_with_resume(email_data, resume):
                self.sent_count += 1
            else:
                self.failed_count += 1
            
            # Track
            update_excel_tracker(recruiter)
        
        # Summary
        print("\n" + "="*60)
        print("✅ OUTREACH COMPLETE")
        print("="*60)
        print(f"📊 Summary:")
        print(f"  ✅ Sent: {self.sent_count}")
        print(f"  ❌ Failed: {self.failed_count}")
        print(f"  📝 Tracked in Excel: recruiter_tracking.xlsx")
        print(f"⏰ Completed at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("="*60 + "\n")

if __name__ == "__main__":
    automation = OutreachAutomation()
    automation.run()
