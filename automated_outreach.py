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
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import base64
import re
from docx import Document
from docx.shared import Pt, Inches
import openpyxl
from openpyxl.utils import get_column_letter

# Configuration
FIRECRAWL_API_KEY = os.getenv('FIRECRAWL_API_KEY')
GMAIL_CREDENTIALS = os.getenv('GMAIL_CREDENTIALS_JSON')
RESUME_TEMPLATE_PATH = os.getenv('RESUME_TEMPLATE_PATH', 'resume_template.txt')

# LinkedIn search queries for India-based recruiters
LINKEDIN_SEARCH_QUERIES = [
    'site:linkedin.com "hiring" "Java developer" India last 24 hours author_role:"Hiring Manager"',
    'site:linkedin.com "recruiting" "Java" India last 24 hours author_role:"Talent Acquisition"',
    'site:linkedin.com "job opening" "Java backend" India last 24 hours author_role:"HR"',
    'site:linkedin.com "we are hiring" "Spring Boot" India last 24 hours',
]

# Your details
ANUJ_NAME = "Anuj Singh"
ANUJ_EMAIL = os.getenv('ANUJ_EMAIL', 'Anuj04004@gmail.com')
ANUJ_PHONE = os.getenv('ANUJ_PHONE', '+91-9528034629')
ANUJ_LINKEDIN = "linkedin.com/in/anujsingh-java"

class OutreachAutomation:
    def __init__(self):
        self.gmail_service = self.setup_gmail()
        self.recruiters = []
        self.sent_count = 0
        self.failed_count = 0
        self.excel_tracker = "recruiter_tracking.xlsx"
        
    def setup_gmail(self):
        """Setup Gmail API connection"""
        try:
            creds_json = json.loads(GMAIL_CREDENTIALS)
            creds = Credentials.from_authorized_user_info(creds_json)
            return build('gmail', 'v1', credentials=creds)
        except Exception as e:
            print(f"❌ Gmail setup failed: {e}")
            return None
    
    def scrape_linkedin_recruiters(self):
        """Scrape LinkedIn for India-based recruiters posting in last 24 hours"""
        print("🔍 Scraping LinkedIn for recruiters...")
        recruiters = []
        
        for query in LINKEDIN_SEARCH_QUERIES:
            try:
                # Use Firecrawl to scrape
                response = requests.post(
                    'https://api.firecrawl.dev/v1/extract',
                    headers={'Authorization': f'Bearer {FIRECRAWL_API_KEY}'},
                    json={
                        'url': f'https://www.google.com/search?q={query}',
                        'extractionSchema': {
                            'type': 'object',
                            'properties': {
                                'recruiter_name': {'type': 'string'},
                                'company': {'type': 'string'},
                                'email': {'type': 'string'},
                                'linkedin_url': {'type': 'string'},
                                'job_description': {'type': 'string'},
                                'post_url': {'type': 'string'},
                                'location': {'type': 'string'}
                            }
                        }
                    }
                )
                
                if response.status_code == 200:
                    data = response.json()
                    if 'data' in data:
                        recruiters.extend(data['data'])
                        print(f"✅ Found {len(data['data'])} recruiters from query")
                        
            except Exception as e:
                print(f"⚠️ Error scraping query: {e}")
        
        # Filter: India only, no duplicates
        self.recruiters = self.filter_india_recruiters(recruiters)
        print(f"✅ Total recruiters (India, deduplicated): {len(self.recruiters)}")
        return self.recruiters
    
    def filter_india_recruiters(self, recruiters):
        """Filter for India-only, remove duplicates"""
        india_locations = [
            'Bengaluru', 'Noida', 'Delhi', 'Mumbai', 'Pune', 'Hyderabad',
            'Bangalore', 'New Delhi', 'NCR', 'Gurgaon', 'Chandigarh'
        ]
        
        filtered = []
        seen_emails = set()
        
        for recruiter in recruiters:
            # Check location
            location = recruiter.get('location', '').lower()
            if not any(city.lower() in location for city in india_locations):
                continue
            
            # Check email
            email = recruiter.get('email', '').strip()
            if not email or email in seen_emails:
                continue
            
            seen_emails.add(email)
            filtered.append(recruiter)
        
        return filtered[:50]  # Max 50 per day
    
    def tailor_resume(self, recruiter_data):
        """Tailor resume to match job description"""
        jd = recruiter_data.get('job_description', '')
        company = recruiter_data.get('company', 'Company')
        
        if not jd:
            print(f"  ℹ️ No JD for {company}, using standard resume")
            return self.get_standard_resume()
        
        print(f"  ✏️ Tailoring resume for {company}...")
        
        # Extract keywords from JD
        keywords = self.extract_keywords(jd)
        
        # Reorder resume based on keywords
        tailored = self.reorder_resume(keywords)
        
        return tailored
    
    def extract_keywords(self, jd):
        """Extract important keywords from job description"""
        keywords = []
        jd_lower = jd.lower()
        
        # Tech skills to look for
        tech_keywords = [
            'kafka', 'redis', 'microservice', 'aws', 'azure', 'kubernetes',
            'docker', 'spring boot', 'java 17', 'java 21', 'postgresql',
            'mongodb', 'system design', 'rest api', 'ci/cd', 'jenkins',
            'github actions', 'distributed system', 'event driven'
        ]
        
        for keyword in tech_keywords:
            if keyword in jd_lower:
                keywords.append(keyword)
        
        return keywords[:5]  # Top 5 keywords
    
    def reorder_resume(self, keywords):
        """Reorder resume to prioritize matching keywords"""
        resume = self.get_standard_resume()
        
        # Simple reordering: move matching skills to front
        if 'kafka' in keywords:
            resume = resume.replace(
                'Databases & Messaging: PostgreSQL',
                'Databases & Messaging: Apache Kafka, Kafka Streams, Redis, PostgreSQL'
            )
        
        if 'kubernetes' in keywords or 'docker' in keywords:
            resume = resume.replace(
                'Cloud & DevOps:',
                'Cloud & DevOps: Kubernetes, Docker,'
            )
        
        if 'aws' in keywords:
            resume = resume.replace(
                'AWS (EC2',
                'AWS (Lambda, EC2'
            )
        
        return resume
    
    def get_standard_resume(self):
        """Load standard resume template"""
        with open(RESUME_TEMPLATE_PATH, 'r') as f:
            return f.read()
    
    def generate_email(self, recruiter_data, resume):
        """Generate personalized email"""
        name = recruiter_data.get('recruiter_name', 'Hiring Manager')
        company = recruiter_data.get('company', 'Company')
        jd = recruiter_data.get('job_description', '')
        
        # Extract relevant detail from JD
        detail = jd.split('.')[0] if jd else 'for a Java position'
        
        email_body = f"""Hi {name},

I came across your recent post about hiring at {company}, and I'm genuinely interested in the opportunity.

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
{ANUJ_EMAIL}
{ANUJ_PHONE}
LinkedIn: {ANUJ_LINKEDIN}"""
        
        subject = f"Java Developer | 4 Years | Spring Boot | Immediate Joiner"
        
        return {
            'subject': subject,
            'body': email_body,
            'to': recruiter_data.get('email')
        }
    
    def send_email_with_resume(self, email_data, resume):
        """Send email with resume via Gmail"""
        if not self.gmail_service:
            print(f"❌ Gmail not configured, skipping email to {email_data['to']}")
            return False
        
        try:
            # Create message
            message = MIMEMultipart()
            message['to'] = email_data['to']
            message['subject'] = email_data['subject']
            
            # Body
            message.attach(MIMEText(email_data['body'], 'plain'))
            
            # Attach resume as PDF (simulated as text for now)
            resume_part = MIMEBase('application', 'octet-stream')
            resume_part.set_payload(resume.encode())
            encoders.encode_base64(resume_part)
            resume_part.add_header(
                'Content-Disposition',
                'attachment; filename= resume.txt'
            )
            message.attach(resume_part)
            
            # Send
            raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
            self.gmail_service.users().messages().send(
                userId='me',
                body={'raw': raw_message}
            ).execute()
            
            print(f"✅ Email sent to {email_data['to']}")
            self.sent_count += 1
            return True
            
        except Exception as e:
            print(f"❌ Failed to send to {email_data['to']}: {e}")
            self.failed_count += 1
            return False
    
    def update_excel_tracker(self, recruiter_data):
        """Update Excel tracking sheet"""
        try:
            # Load or create workbook
            if os.path.exists(self.excel_tracker):
                wb = openpyxl.load_workbook(self.excel_tracker)
            else:
                wb = openpyxl.Workbook()
            
            ws = wb.active
            
            # Add row
            row = ws.max_row + 1
            ws[f'A{row}'] = recruiter_data.get('recruiter_name')
            ws[f'B{row}'] = recruiter_data.get('email')
            ws[f'C{row}'] = recruiter_data.get('company')
            ws[f'D{row}'] = recruiter_data.get('job_description', '')[:50]
            ws[f'E{row}'] = 'India'
            ws[f'F{row}'] = datetime.now().strftime('%Y-%m-%d')
            ws[f'G{row}'] = 'Sent'
            ws[f'H{row}'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            wb.save(self.excel_tracker)
            
        except Exception as e:
            print(f"⚠️ Excel update error: {e}")
    
    def run(self):
        """Main execution"""
        print("\n" + "="*60)
        print("🚀 AUTOMATED RECRUITER OUTREACH - STARTED")
        print("="*60)
        print(f"⏰ Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        
        # Step 1: Scrape
        self.scrape_linkedin_recruiters()
        
        if not self.recruiters:
            print("❌ No recruiters found")
            return
        
        # Step 2: Generate & Send emails
        print(f"\n📧 Generating and sending {len(self.recruiters)} emails...\n")
        
        for i, recruiter in enumerate(self.recruiters, 1):
            print(f"\n[{i}/{len(self.recruiters)}] Processing: {recruiter.get('recruiter_name')}")
            
            # Tailor resume
            resume = self.tailor_resume(recruiter)
            
            # Generate email
            email_data = self.generate_email(recruiter, resume)
            
            # Send
            self.send_email_with_resume(email_data, resume)
            
            # Track
            self.update_excel_tracker(recruiter)
        
        # Summary
        print("\n" + "="*60)
        print("✅ OUTREACH COMPLETE")
        print("="*60)
        print(f"📊 Summary:")
        print(f"  ✅ Sent: {self.sent_count}")
        print(f"  ❌ Failed: {self.failed_count}")
        print(f"  📝 Tracked in Excel: {self.excel_tracker}")
        print(f"⏰ Completed at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("="*60 + "\n")

if __name__ == "__main__":
    automation = OutreachAutomation()
    automation.run()
