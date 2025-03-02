import sqlite3
import requests
import smtplib
import os
import logging
import time
import re
import unicodedata
import json
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta
from dotenv import load_dotenv
from bs4 import BeautifulSoup
import nltk
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.corpus import stopwords
import langdetect  # For language detection
from collections import Counter
import hashlib

# Load environment variables from .env file
load_dotenv()

# Configure logging
logging.basicConfig(
    filename='office_tracker.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Example companies to track - these are just examples, the system will detect any company
EXAMPLE_COMPANIES = [
    "PwC", "KPMG", "Deloitte", "EY", "Lamda Development", 
    "Dimand", "Prodea", "Noval Property"
]

# Company type identifiers - used to identify companies in text
COMPANY_IDENTIFIERS = {
    "en": [
        "Inc", "LLC", "Ltd", "Limited", "Corp", "Corporation", "Co", "Company", 
        "Group", "Holdings", "Enterprises", "Ventures", "Capital", "Partners",
        "Properties", "Real Estate", "Development", "Investments"
    ],
    "el": [
        "ΑΕ", "Α.Ε.", "Α.Ε", "ΕΠΕ", "Ε.Π.Ε.", "Ε.Π.Ε", "ΟΕ", "Ο.Ε.", "ΙΚΕ", "Ι.Κ.Ε.",
        "ΑΕΒΕ", "Α.Ε.Β.Ε.", "ΑΒΕΕ", "Α.Β.Ε.Ε.", "ΑΞΤΕ", "Α.Ξ.Τ.Ε.",
        "Όμιλος", "Εταιρεία", "Εταιρεια", "Ανάπτυξη", "Αναπτυξη", "Ακίνητα", "Ακινητα", 
        "Επενδύσεις", "Επενδυσεις", "Κατασκευαστική", "Κατασκευαστικη"
    ]
}

# Step 1: Create and setup the database
def create_database():
    conn = sqlite3.connect("office_projects.db")
    cursor = conn.cursor()
    
    # Projects table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS projects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_name TEXT,
            location TEXT,
            description TEXT,
            source_url TEXT,
            source_title TEXT,
            relevance_score REAL,
            project_type TEXT,
            estimated_size TEXT,
            date_reported TEXT,
            date_added TEXT,
            content_hash TEXT,
            is_sent BOOLEAN DEFAULT 0
        )
    ''')
    
    # Companies table to track discovered companies
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS companies (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE,
            mention_count INTEGER DEFAULT 1,
            last_seen_date TEXT,
            relevance_score REAL DEFAULT 0
        )
    ''')
    
    # Locations table to track office locations
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS locations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE,
            mention_count INTEGER DEFAULT 1,
            last_seen_date TEXT
        )
    ''')
    
    # Analytics table for tracking metrics
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS analytics (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT,
            searches_performed INTEGER,
            results_found INTEGER,
            relevant_results INTEGER,
            emails_sent INTEGER,
            performance_metrics TEXT
        )
    ''')
    
    # Create an index on source_url for faster duplicate checking
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_source_url ON projects(source_url)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_content_hash ON projects(content_hash)')
    
    # Search queries log table to track what's been searched
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS search_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            query TEXT,
            language TEXT,
            date_searched TEXT,
            results_count INTEGER
        )
    ''')
    
    conn.commit()
    conn.close()
    logging.info("Database initialized successfully")


# Step 2: Search the web using multiple search strategies
def search_web(api_key, cse_id):
    all_results = []
    searches_performed = 0
    search_start_time = time.time()
    
    # Generate search queries in both English and Greek
    search_queries = [
        # English queries
        "office projects",
        "new office development",
        "commercial real estate acquisition",
        "office relocation",
        "office building purchase",
        "new headquarters",
        "corporate office move",
        "office space leasing",
        "business district development",
        # Greek queries
        "γραφεία έργα",
        "νέα γραφεία",
        "επαγγελματικά ακίνητα",
        "ανάπτυξη γραφείων",
        "αγορά κτιρίου γραφείων",
        "μετεγκατάσταση γραφείων",
        "επαγγελματική στέγη",
        "νέα έδρα εταιρείας",
        "επένδυση ακινήτων γραφείων",
        "εμπορικό ακίνητο"
    ]
    
    # Add location-specific queries
    locations = ["Athens", "Thessaloniki", "Patras", "Heraklion", "Piraeus", 
                "Αθήνα", "Θεσσαλονίκη", "Πάτρα", "Ηράκλειο", "Πειραιάς"]
    
    for base_query in search_queries[:6]:  # Use only main queries for location combinations
        for location in locations[:5]:  # Filter to main locations
            if "Greece" not in base_query:
                search_queries.append(f"{base_query} in {location}, Greece")
    
    for base_query in search_queries[9:15]:  # Use only main Greek queries
        for location in locations[5:]:  # Filter to main Greek locations
            search_queries.append(f"{base_query} {location}")
    
    # Add queries for specific project types
    project_types = [
        "office renovation", "office expansion", "office campus", 
        "tech hub", "corporate campus", "innovation center",
        "ανακαίνιση γραφείων", "επέκταση γραφείων", "κέντρο καινοτομίας"
    ]
    
    for project_type in project_types:
        search_queries.append(f"{project_type} Greece")
    
    # Add company-specific queries for example companies
    for company in EXAMPLE_COMPANIES:
        search_queries.append(f"{company} new office Greece")
    
    # Get previously successful queries from database
    conn = sqlite3.connect("office_projects.db")
    cursor = conn.cursor()
    cutoff_date = (datetime.now() - timedelta(days=7)).strftime("%Y-%m-%d")
    
    cursor.execute('''
        SELECT query FROM search_log 
        WHERE date_searched > ? AND results_count > 0
        GROUP BY query 
        ORDER BY SUM(results_count) DESC 
        LIMIT 5
    ''', (cutoff_date,))
    
    successful_queries = [row[0] for row in cursor.fetchall()]
    conn.close()
    
    # Add successful queries from past (with priority)
    search_queries = successful_queries + [q for q in search_queries if q not in successful_queries]
    
    # Limit queries to avoid too many API calls
    max_queries = int(os.environ.get("MAX_SEARCH_QUERIES", "30"))
    search_queries = search_queries[:max_queries]
    
    # Search with each query
    for query in search_queries:
        try:
            # Detect query language
            query_lang = detect_language(query)
            
            url = f"https://www.googleapis.com/customsearch/v1?q={query}&cx={cse_id}&key={api_key}&dateRestrict=d7"
            response = requests.get(url, timeout=30)
            searches_performed += 1
            
            if response.status_code == 200:
                results = response.json().get('items', [])
                
                # Log the search
                conn = sqlite3.connect("office_projects.db")
                cursor = conn.cursor()
                cursor.execute(
                    "INSERT INTO search_log (query, language, date_searched, results_count) VALUES (?, ?, ?, ?)",
                    (query, query_lang, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), len(results))
                )
                conn.commit()
                conn.close()
                
                # Process and add results
                for item in results:
                    result = {
                        'title': item.get('title', ''),
                        'link': item.get('link', ''),
                        'snippet': item.get('snippet', ''),
                        'query': query,
                        'query_language': query_lang
                    }
                    
                    # Create a hash for duplicate detection
                    content_hash = hashlib.md5(f"{result['title']}|{result['snippet']}".encode('utf-8')).hexdigest()
                    result['content_hash'] = content_hash
                    
                    # Skip if we've already seen this result
                    if any(r.get('content_hash') == content_hash for r in all_results):
                        continue
                    
                    # Fetch full content if possible
                    try:
                        page_resp = requests.get(result['link'], timeout=20)
                        if page_resp.status_code == 200:
                            soup = BeautifulSoup(page_resp.text, 'html.parser')
                            # Get the main content if possible
                            main_content = soup.find('main') or soup.find('article') or soup.body
                            if main_content:
                                result['full_content'] = main_content.get_text(strip=True)[:10000]  # Limit content length
                            else:
                                result['full_content'] = soup.get_text(strip=True)[:5000]  # Limit content length
                            
                            # Extract publication date if available
                            pub_date = None
                            date_meta = soup.find('meta', {'property': 'article:published_time'}) or \
                                      soup.find('meta', {'name': 'pubdate'}) or \
                                      soup.find('meta', {'name': 'date'})
                            
                            if date_meta and 'content' in date_meta.attrs:
                                pub_date = date_meta['content'][:10]  # Get YYYY-MM-DD portion
                            
                            if pub_date:
                                result['publication_date'] = pub_date
                    except Exception as e:
                        logging.warning(f"Could not fetch full content for {result['link']}: {e}")
                        result['full_content'] = result['snippet']
                    
                    all_results.append(result)
                
                # Avoid hitting rate limits
                time.sleep(2)
            else:
                logging.error(f"Error with query '{query}': {response.status_code} - {response.text}")
        except Exception as e:
            logging.error(f"Exception during search for query '{query}': {e}")
            continue
    
    # Record search performance metrics
    search_duration = time.time() - search_start_time
    conn = sqlite3.connect("office_projects.db")
    cursor = conn.cursor()
    
    performance_metrics = json.dumps({
        'search_duration_seconds': search_duration,
        'avg_time_per_query': search_duration / max(searches_performed, 1),
        'queries_performed': searches_performed,
        'results_per_query': len(all_results) / max(searches_performed, 1)
    })
    
    cursor.execute('''
        INSERT INTO analytics (date, searches_performed, results_found, performance_metrics)
        VALUES (?, ?, ?, ?)
    ''', (datetime.now().strftime("%Y-%m-%d"), searches_performed, len(all_results), performance_metrics))
    
    conn.commit()
    conn.close()
    
    return all_results

# Extract company names from text
def extract_company_names(text, language="en"):
    """
    Extract potential company names from text content.
    Uses patterns and company identifiers to find potential company mentions.
    """
    if not text:
        return []
    
    companies = []
    
    # Get appropriate company identifiers based on language
    identifiers = COMPANY_IDENTIFIERS.get(language, COMPANY_IDENTIFIERS["en"])
    
    # First pass: look for companies with known suffixes
    for identifier in identifiers:
        # Pattern to match: Word(s) + company identifier
        pattern = r'([A-Za-z\u0370-\u03FF\u1F00-\u1FFF][\w\s&\-\'\u0370-\u03FF\u1F00-\u1FFF]{2,50}?\s)' + re.escape(identifier) + r'\b'
        matches = re.finditer(pattern, text)
        
        for match in matches:
            company_name = match.group(1).strip() + identifier
            if 3 < len(company_name) < 60:  # Reasonable length for company name
                companies.append(company_name.strip())
    
    # Second pass: try to extract from common patterns like "Company X announced today"
    if language == "en":
        company_verbs = ["announced", "reported", "unveiled", "launched", "introduced", 
                         "acquired", "purchased", "bought", "leased", "moved", "relocated"]
        
        for verb in company_verbs:
            pattern = r'([A-Z][\w\s&\-\']{2,50}?)\s' + re.escape(verb) + r'\b'
            matches = re.finditer(pattern, text)
            
            for match in matches:
                company_name = match.group(1).strip()
                if 3 < len(company_name) < 60 and company_name not in companies:
                    companies.append(company_name)
    
    # Greek extraction patterns
    if language == "el":
        company_verbs = ["ανακοίνωσε", "παρουσίασε", "απέκτησε", "αγόρασε", "μίσθωσε", "μετεγκαταστάθηκε"]
        
        for verb in company_verbs:
            pattern = r'([Α-ΩΆΈΉΊΌΎΏΪΫ][\w\s&\-\'\u0370-\u03FF\u1F00-\u1FFF]{2,50}?)\s' + re.escape(verb) + r'\b'
            matches = re.finditer(pattern, text)
            
            for match in matches:
                company_name = match.group(1).strip()
                if 3 < len(company_name) < 60 and company_name not in companies:
                    companies.append(company_name)
    
    # Third pass: Look for quoted company names
    if language == "en":
        matches = re.finditer(r'"([^"]{3,60}?)"[\s,\.]+(a|an|the)\s+(?:leading|global|company|firm|developer|investor)', text, re.IGNORECASE)
        for match in matches:
            company_name = match.group(1).strip()
            if company_name not in companies:
                companies.append(company_name)
    
    if language == "el":
        matches = re.finditer(r'"([^"]{3,60}?)"[\s,\.]+(η|ο|το)\s+(?:εταιρεία|εταιρία|όμιλος|επενδυτής|κατασκευαστική)', text, re.IGNORECASE)
        for match in matches:
            company_name = match.group(1).strip()
            if company_name not in companies:
                companies.append(company_name)
    
    # Check for example companies we already know about
    for company in EXAMPLE_COMPANIES:
        if company in text and company not in companies:
            companies.append(company)
    
    # Remove duplicates and short entries
    companies = [c for c in companies if len(c) > 3]
    return list(set(companies))

# Extract location information from text
def extract_locations(text, language="en"):
    """Extract potential location information from text."""
    locations = []
    
    if language == "en":
        # English location patterns
        patterns = [
            r'(?:in|at|near|to|from)\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?),?\s+(?:Greece|Hellas)',
            r'(?:relocating|moved|moving)\s+(?:to|into|in)\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)',
            r'([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)\s+(?:district|area|suburb|region|business\s+center)',
            r'new\s+(?:office|headquarters|building)\s+(?:in|at)\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)'
        ]
        
        for pattern in patterns:
            matches = re.finditer(pattern, text)
            for match in matches:
                location = match.group(1).strip()
                if location and len(location) > 2 and location not in locations:
                    locations.append(location)
    
    if language == "el":
        # Greek location patterns
        patterns = [
            r'(?:στ(?:ην|ον|ο|α)|κοντά\s+στ(?:ην|ον|ο|α))\s+([Α-ΩΆΈΉΊΌΎΏΪΫ][α-ωάέήίόύώϊϋ]+(?:\s+[Α-ΩΆΈΉΊΌΎΏΪΫ][α-ωάέήίόύώϊϋ]+)?)',
            r'(?:μετεγκατάσταση|μετακόμιση)\s+(?:στ(?:ην|ον|ο|α))\s+([Α-ΩΆΈΉΊΌΎΏΪΫ][α-ωάέήίόύώϊϋ]+(?:\s+[Α-ΩΆΈΉΊΌΎΏΪΫ][α-ωάέήίόύώϊϋ]+)?)',
            r'([Α-ΩΆΈΉΊΌΎΏΪΫ][α-ωάέήίόύώϊϋ]+(?:\s+[Α-ΩΆΈΉΊΌΎΏΪΫ][α-ωάέήίόύώϊϋ]+)?)\s+(?:περιοχή|συνοικία|προάστιο)'
        ]
        
        for pattern in patterns:
            matches = re.finditer(pattern, text)
            for match in matches:
                location = match.group(1).strip()
                if location and len(location) > 2 and location not in locations:
                    locations.append(location)
    
    # Check for common Greek cities/areas
    common_locations = [
        "Athens", "Thessaloniki", "Patras", "Heraklion", "Piraeus", "Larissa", "Glyfada", 
        "Marousi", "Chalandri", "Kifissia", "Kallithea", "Palaio Faliro", "Voula", 
        "Αθήνα", "Θεσσαλονίκη", "Πάτρα", "Ηράκλειο", "Πειραιάς", "Λάρισα", "Γλυφάδα",
        "Μαρούσι", "Χαλάνδρι", "Κηφισιά", "Καλλιθέα", "Παλαιό Φάληρο", "Βούλα"
    ]
    
    for location in common_locations:
        if location in text and location not in locations:
            locations.append(location)
    
    return locations

# Extract project type from content
def extract_project_type(content, language="en"):
    """Determine the type of office project based on content analysis."""
    content_lower = content.lower()
    
    # English project types
    if language == "en":
        if any(term in content_lower for term in ["new office", "new building", "new headquarters", "new hq"]):
            return "New Office"
        
        if any(term in content_lower for term in ["relocation", "relocating", "moving to", "moved to"]):
            return "Relocation"
        
        if any(term in content_lower for term in ["expansion", "expanding", "additional space", "growing"]):
            return "Expansion"
        
        if any(term in content_lower for term in ["renovation", "refurbishment", "remodeling", "upgrading"]):
            return "Renovation"
        
        if any(term in content_lower for term in ["lease", "leasing", "leased", "rental", "renting"]):
            return "Leasing"
        
        if any(term in content_lower for term in ["purchase", "acquisition", "acquired", "bought", "buying"]):
            return "Acquisition"
    
    # Greek project types
    if language == "el":
        if any(term in content_lower for term in ["νέο γραφείο", "νέα γραφεία", "νέο κτίριο", "νέα έδρα"]):
            return "New Office"
        
        if any(term in content_lower for term in ["μετεγκατάσταση", "μετακόμιση", "μεταφορά γραφείων"]):
            return "Relocation"
        
        if any(term in content_lower for term in ["επέκταση", "επέκταση γραφείων", "επιπλέον χώρος"]):
            return "Expansion"
        
        if any(term in content_lower for term in ["ανακαίνιση", "ανακαίνιση γραφείων", "αναβάθμιση"]):
            return "Renovation"
        
        if any(term in content_lower for term in ["μίσθωση", "ενοικίαση", "μισθώνει", "ενοικιάζει"]):
            return "Leasing"
        
        if any(term in content_lower for term in ["αγορά", "εξαγορά", "απόκτηση", "αγοράζει"]):
            return "Acquisition"
    
    return "Office Project"  # Default
    
    # Try to download NLTK resources if not already available
    try:
        nltk.data.find('tokenizers/punkt')
    except LookupError:
        nltk.download('punkt')
    
    try:
        nltk.data.find('corpora/stopwords')
    except LookupError:
        nltk.download('stopwords')
    
    # Get both English and Greek stopwords if available
    stop_words = set(stopwords.words('english'))
    try:
        # Add Greek stopwords if available
        greek_stop_words = set(stopwords.words('greek'))
        stop_words = stop_words.union(greek_stop_words)
    except:
        # If Greek not available, define basic Greek stopwords
        greek_stop_words = {
            'ο', 'η', 'το', 'οι', 'τα', 'του', 'της', 'των', 'τον', 'την', 'και',
            'κι', 'κ', 'με', 'σε', 'από', 'για', 'προς', 'παρά', 'αντί', 'μέχρι',
            'ως', 'πως', 'ότι', 'είναι', 'ήταν', 'θα', 'να', 'δεν', 'μη', 'μην'
        }
        stop_words = stop_words.union(greek_stop_words)
    
    for result in results:
        # Initial score based on the presence of company names
        score = 0
        company_found = None
        
        # Identify company mentions
        for company in COMPANIES_TO_TRACK:
            company_lower = company.lower()
            if company_lower in result['title'].lower() or company_lower in result['snippet'].lower() or company_lower in result.get('full_content', '').lower():
                score += 15
                company_found = company
                break
        
        # Check for office-related keywords
        content_to_check = f"{result['title']} {result['snippet']} {result.get('full_content', '')}"
        for keyword in office_keywords:
            if keyword.lower() in content_to_check.lower():
                score += 10
        
        # Analyze content with NLP to extract relevant information
        try:
            # Tokenize and analyze the text - handle both English and Greek
            # Try to detect if content is primarily Greek
            is_greek = False
            greek_chars = set('αβγδεζηθικλμνξοπρστυφχψωάέήίόύώϊϋΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩΆΈΉΊΌΎΏΪΫ')
            sample_text = content_to_check[:1000].lower()  # Take a sample for analysis
            greek_char_count = sum(1 for c in sample_text if c in greek_chars)
            
            if greek_char_count > len(sample_text) * 0.3:  # If more than 30% Greek characters
                is_greek = True
                
            # Tokenize text
            tokens = word_tokenize(content_to_check.lower())
            filtered_tokens = [w for w in tokens if w not in stop_words]
            
            # Add additional score for Greek content when searching specifically for Greek content
            if is_greek and any(q for q in result.get('query', '').lower() if 'ελλάδα' in q.lower() or 'γραφεία' in q.lower()):
                score += 5
            
            # Count occurrences of office-related terms
            office_term_count = sum(1 for token in filtered_tokens if any(keyword.lower() in token for keyword in office_keywords))
            score += min(office_term_count * 2, 20)  # Cap at 20 points
            
            # Extract potential location information (in both English and Greek)
            # English location pattern
            eng_location_pattern = r'\b(?:in|at|near)\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?),?\s+(?:Greece|Hellas)\b'
            eng_location_matches = re.findall(eng_location_pattern, result['full_content'] if 'full_content' in result else result['snippet'])
            
            # Greek location pattern
            greek_location_pattern = r'\b(?:στ(?:ην|ον|ο|α)|κοντά στ(?:ην|ον|ο|α))\s+([Α-ΩΆΈΉΊΌΎΏΪΫ][α-ωάέήίόύώϊϋ]+(?:\s+[Α-ΩΆΈΉΊΌΎΏΪΫ][α-ωάέήίόύώϊϋ]+)?),?\s+(?:Ελλάδα|Ελλάδας)\b'
            greek_location_matches = re.findall(greek_location_pattern, result['full_content'] if 'full_content' in result else result['snippet'])
            
            # Combined matches
            location_matches = eng_location_matches + greek_location_matches
            location = location_matches[0] if location_matches else "Greece"
            
            # Extract potential size information (in both English and Greek)
            size_pattern = r'\b(\d[\d,.]*)\s*(?:sq\.?m|square meters|m²|sqm|τ\.?μ\.?|τετραγωνικά μέτρα|τετραγωνικών μέτρων)\b'
            size_matches = re.findall(size_pattern, result['full_content'] if 'full_content' in result else result['snippet'])
            
            if size_matches:
                score += 10
            
        except Exception as e:
            logging.warning(f"Error during NLP processing: {e}")
        
        # If we have a reasonable relevance score, add it to our results
        if score >= 30:
            result['relevance_score'] = score
            result['extracted_company'] = company_found if company_found else "Unknown"
            result['extracted_location'] = location if 'location' in locals() else "Unknown"
            relevant_results.append(result)
    
    return relevant_results

# Step 4: Save relevant results to the database, avoiding duplicates
def save_to_database(projects):
    conn = sqlite3.connect("office_projects.db")
    cursor = conn.cursor()
    
    new_projects_count = 0
    
    for project in projects:
        # Check if URL already exists
        cursor.execute('SELECT id FROM projects WHERE source_url = ?', (project['link'],))
        existing_url = cursor.fetchone()
        
        # Check if content hash exists (for duplicates with different URLs)
        cursor.execute('SELECT id FROM projects WHERE content_hash = ?', (project.get('content_hash', ''),))
        existing_hash = cursor.fetchone()
        
        if existing_url or existing_hash:
            logging.info(f"Skipping duplicate content: {project['link']}")
            continue
        
        # Check for similar content to avoid duplicates with different URLs
        snippet_words = project['snippet'].lower().split()
        if len(snippet_words) > 10:  # Only check substantial snippets
            # Create a simplified version of the snippet for comparison
            simple_snippet = ' '.join(snippet_words[:20])  # First 20 words
            
            cursor.execute('SELECT id FROM projects WHERE description LIKE ?', (f"%{simple_snippet}%",))
            similar = cursor.fetchone()
            
            if similar:
                logging.info(f"Skipping similar content: {project['link']}")
                continue
        
        # If we get here, it's a new project - insert it
        try:
            cursor.execute('''
                INSERT INTO projects (
                    company_name, 
                    location, 
                    description, 
                    source_url, 
                    source_title,
                    relevance_score,
                    project_type,
                    estimated_size,
                    date_reported,
                    date_added,
                    content_hash
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                project.get('extracted_company', 'Unknown'),
                project.get('extracted_location', 'Greece'),
                project['snippet'],
                project['link'],
                project['title'],
                project.get('relevance_score', 0),
                project.get('project_type', 'Office Project'),
                project.get('estimated_size', 'Unknown'),
                project.get('date_reported', datetime.now().strftime("%Y-%m-%d")),
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                project.get('content_hash', '')
            ))
            new_projects_count += 1
        except Exception as e:
            logging.error(f"Error saving project to database: {e}")
    
    # Update analytics 
    cursor.execute('''
        UPDATE analytics 
        SET relevant_results = ? 
        WHERE date = ?
    ''', (new_projects_count, datetime.now().strftime("%Y-%m-%d")))
    
    if cursor.rowcount == 0:  # No row updated, insert new one
        cursor.execute('''
            INSERT INTO analytics (date, relevant_results)
            VALUES (?, ?)
        ''', (datetime.now().strftime("%Y-%m-%d"), new_projects_count))
    
    conn.commit()
    conn.close()
    
    logging.info(f"Saved {new_projects_count} new projects to database")
    return new_projects_count

# Step 5: Send email notifications with improved formatting
def send_email_notification(receiver_email):
    # Get environment variables for email configuration
    sender_email = os.environ.get("EMAIL_USERNAME")
    sender_password = os.environ.get("EMAIL_PASSWORD")
    smtp_server = os.environ.get("SMTP_SERVER", "smtp.gmail.com")
    smtp_port = int(os.environ.get("SMTP_PORT", "587"))
    
    # Email language settings
    email_language = os.environ.get("EMAIL_LANGUAGE", "en").lower()  # Default to English, can be "el" for Greek
    
    if not all([sender_email, sender_password]):
        logging.error("Email credentials not found in environment variables")
        return False
    
    # Get unsent projects from database
    conn = sqlite3.connect("office_projects.db")
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute('''
        SELECT * FROM projects
        WHERE is_sent = 0
        ORDER BY relevance_score DESC, date_added DESC
    ''')
    unsent_projects = cursor.fetchall()
    
    if not unsent_projects:
        logging.info("No new projects to send notification about")
        return False
    
    # Create email
    msg = MIMEMultipart('alternative')
    msg['From'] = sender_email
    msg['To'] = receiver_email
    # Bilingual email subject and intro text
    if email_language == "el":
        msg['Subject'] = f"Ενημέρωση Έργων Γραφείων - {datetime.now().strftime('%Y-%m-%d')}"
        text_intro = "Παρακάτω θα βρείτε τα τελευταία έργα γραφείων που εντοπίστηκαν:\n\n"
        html_intro = "Παρακάτω θα βρείτε τα τελευταία έργα γραφείων που εντοπίστηκαν:"
        header_text = f"Ενημέρωση Έργων Γραφείων - {datetime.now().strftime('%Y-%m-%d')}"
        location_label = "Τοποθεσία:"
        source_label = "Πηγή:"
        relevance_label = "Βαθμός Συνάφειας:"
        read_more = "Διαβάστε Περισσότερα"
    else:  # Default to English
        msg['Subject'] = f"Office Project Updates - {datetime.now().strftime('%Y-%m-%d')}"
        text_intro = "Here are the latest office projects found:\n\n"
        html_intro = "Here are the latest office projects found:"
        header_text = f"Office Project Updates - {datetime.now().strftime('%Y-%m-%d')}"
        location_label = "Location:"
        source_label = "Source:"
        relevance_label = "Relevance Score:"
        read_more = "Read More"
    
    # Create plain text version
    text_body = text_intro
    for project in unsent_projects:
        text_body += f"- {project['company_name']} in {project['location']}\n"
        text_body += f"  {project['source_title']}\n"
        text_body += f"  {project['description']}\n"
        text_body += f"  Source: {project['source_url']}\n\n"
    
    # Create HTML version
    html_body = f"""
    <html>
    <head>
        <style>
            body {{ font-family: Arial, sans-serif; line-height: 1.6; }}
            .project {{ margin-bottom: 20px; padding: 15px; border: 1px solid #ddd; border-radius: 5px; }}
            .company {{ font-weight: bold; color: #2c3e50; }}
            .location {{ color: #7f8c8d; }}
            .title {{ font-size: 18px; color: #16a085; margin: 5px 0; }}
            .description {{ margin: 10px 0; }}
            .link {{ color: #3498db; }}
            .relevance {{ font-size: 12px; color: #95a5a6; text-align: right; }}
        </style>
        <meta charset="UTF-8">
    </head>
    <body>
        <h2>{header_text}</h2>
        <p>{html_intro}</p>
    """
    
    for project in unsent_projects:
        html_body += f"""
        <div class="project">
            <div class="company">{project['company_name']}</div>
            <div class="location">{location_label} {project['location']}</div>
            <div class="title">{project['source_title']}</div>
            <div class="description">{project['description']}</div>
            <div><a href="{project['source_url']}" class="link">{read_more}</a></div>
            <div class="relevance">{relevance_label} {project['relevance_score']}</div>
        </div>
        """
    
    html_body += """
    </body>
    </html>
    """
    
    # Attach both versions to the email
    part1 = MIMEText(text_body, 'plain')
    part2 = MIMEText(html_body, 'html')
    msg.attach(part1)
    msg.attach(part2)
    
    # Send email
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
        
        # Mark projects as sent
        for project in unsent_projects:
            cursor.execute('UPDATE projects SET is_sent = 1 WHERE id = ?', (project['id'],))
        
        conn.commit()
        logging.info(f"Email sent successfully with {len(unsent_projects)} projects!")
        success = True
    except Exception as e:
        logging.error(f"Failed to send email: {e}")
        success = False
    
    conn.close()
    return success

# Main function
# Add analytics and reporting function
def generate_analytics_report():
    """Generate an analytics report of project tracking trends"""
    conn = sqlite3.connect("office_projects.db")
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    report = {
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "summary": {},
        "top_companies": [],
        "top_locations": [],
        "project_types": {},
        "recent_trends": {}
    }
    
    # Summary statistics
    cursor.execute("""
        SELECT 
            COUNT(*) as total_projects,
            COUNT(DISTINCT company_name) as unique_companies,
            COUNT(DISTINCT location) as unique_locations,
            AVG(relevance_score) as avg_relevance,
            MAX(date_added) as latest_project
        FROM projects
    """)
    summary = cursor.fetchone()
    report["summary"] = dict(summary)
    
    # Top companies by mention count
    cursor.execute("""
        SELECT name, mention_count, last_seen_date
        FROM companies
        ORDER BY mention_count DESC
        LIMIT 15
    """)
    report["top_companies"] = [dict(row) for row in cursor.fetchall()]
    
    # Top locations
    cursor.execute("""
        SELECT name, mention_count, last_seen_date
        FROM locations
        ORDER BY mention_count DESC
        LIMIT 10
    """)
    report["top_locations"] = [dict(row) for row in cursor.fetchall()]
    
    # Project types distribution
    cursor.execute("""
        SELECT project_type, COUNT(*) as count
        FROM projects
        GROUP BY project_type
        ORDER BY count DESC
    """)
    for row in cursor.fetchall():
        report["project_types"][row["project_type"] or "Unknown"] = row["count"]
    
    # Recent trends (last 30 days)
    thirty_days_ago = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d")
    cursor.execute("""
        SELECT 
            date(date_added) as day,
            COUNT(*) as project_count
        FROM projects
        WHERE date_added > ?
        GROUP BY day
        ORDER BY day
    """, (thirty_days_ago,))
    
    for row in cursor.fetchall():
        report["recent_trends"][row["day"]] = row["project_count"]
    
    conn.close()
    return report

# Function to clean up database (remove old search logs, optimize)
def perform_maintenance():
    """Perform database maintenance tasks"""
    conn = sqlite3.connect("office_projects.db")
    cursor = conn.cursor()
    
    # Remove old search logs (older than 30 days)
    thirty_days_ago = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d")
    cursor.execute("DELETE FROM search_log WHERE date_searched < ?", (thirty_days_ago,))
    deleted_logs = cursor.rowcount
    
    # Optimize database
    cursor.execute("VACUUM")
    
    # Update company relevance scores based on mention frequency
    cursor.execute("""
        UPDATE companies
        SET relevance_score = mention_count * 
            CASE 
                WHEN julianday('now') - julianday(last_seen_date) < 30 THEN 2
                WHEN julianday('now') - julianday(last_seen_date) < 90 THEN 1
                ELSE 0.5
            END
    """)
    
    conn.commit()
    conn.close()
    
    logging.info(f"Maintenance complete: removed {deleted_logs} old search logs, optimized database")

# Main function
def main():
    try:
        logging.info("Starting office project search...")
        
        # Create database if it doesn't exist
        create_database()
        
        # Get configuration from environment variables
        api_key = os.environ.get("GOOGLE_API_KEY")
        cse_id = os.environ.get("GOOGLE_CSE_ID")
        receiver_email = os.environ.get("RECEIVER_EMAIL", "panos.bompolas@inmind.com.gr")
        
        # Install required NLTK packages if needed
        try:
            nltk.download('punkt')
            nltk.download('stopwords')
        except:
            logging.warning("Could not download NLTK resources. Will proceed with basic functionality.")
        
        if not all([api_key, cse_id]):
            logging.error("API credentials not found in environment variables")
            return
        
        # Perform database maintenance once a week (on Sundays)
        if datetime.now().weekday() == 6:  # Sunday
            perform_maintenance()
        
        # Search for office projects
        results = search_web(api_key, cse_id)
        logging.info(f"Found {len(results)} initial results")
        
        if results:
            # Evaluate relevance and extract entities
            relevant_results = evaluate_relevance(results)
            logging.info(f"Filtered to {len(relevant_results)} relevant results")
            
            # Save to database
            new_projects_count = save_to_database(relevant_results)
            
            # Send email if there are new projects
            if new_projects_count > 0:
                send_email_notification(receiver_email)
            else:
                logging.info("No new projects to notify about")
        else:
            logging.info("No results found in today's search")
        
        # Generate and log analytics once a week
        if datetime.now().weekday() == 0:  # Monday
            analytics = generate_analytics_report()
            logging.info(f"Weekly Analytics: {json.dumps(analytics['summary'])}")
            
            # Send analytics email if configured
            if os.environ.get("SEND_ANALYTICS", "false").lower() == "true":
                send_analytics_email(receiver_email, analytics)
    
    except Exception as e:
        logging.error(f"Error in main function: {e}")

# Function to send analytics email
def send_analytics_email(receiver_email, analytics):
    """Send an email with analytics information"""
    # Get email settings
    sender_email = os.environ.get("EMAIL_USERNAME")
    sender_password = os.environ.get("EMAIL_PASSWORD")
    smtp_server = os.environ.get("SMTP_SERVER", "smtp.gmail.com")
    smtp_port = int(os.environ.get("SMTP_PORT", "587"))
    
    if not all([sender_email, sender_password]):
        logging.error("Email credentials not found in environment variables")
        return False
    
    # Create email
    msg = MIMEMultipart('alternative')
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = f"Office Project Tracker - Weekly Analytics Report"
    
    # Create HTML content for analytics
    html_body = f"""
    <html>
    <head>
        <style>
            body {{ font-family: Arial, sans-serif; line-height: 1.6; }}
            .container {{ max-width: 800px; margin: 0 auto; padding: 20px; }}
            .header {{ background-color: #2c3e50; color: white; padding: 15px; text-align: center; }}
            .section {{ margin: 20px 0; }}
            h2 {{ color: #2c3e50; border-bottom: 1px solid #eee; padding-bottom: 5px; }}
            table {{ width: 100%; border-collapse: collapse; }}
            th, td {{ padding: 10px; text-align: left; border-bottom: 1px solid #ddd; }}
            th {{ background-color: #f5f5f5; }}
            .card {{ background-color: #f9f9f9; border-radius: 5px; padding: 15px; margin-bottom: 20px; }}
            .highlight {{ color: #e74c3c; font-weight: bold; }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>Office Project Tracker - Weekly Analytics</h1>
                <p>Generated on {analytics['generated_at']}</p>
            </div>
            
            <div class="section">
                <h2>Summary</h2>
                <div class="card">
                    <p><strong>Total Projects:</strong> {analytics['summary']['total_projects']}</p>
                    <p><strong>Unique Companies:</strong> {analytics['summary']['unique_companies']}</p>
                    <p><strong>Unique Locations:</strong> {analytics['summary']['unique_locations']}</p>
                    <p><strong>Average Relevance Score:</strong> {analytics['summary']['avg_relevance']:.2f}</p>
                    <p><strong>Latest Project Added:</strong> {analytics['summary']['latest_project']}</p>
                </div>
            </div>
            
            <div class="section">
                <h2>Top Companies</h2>
                <table>
                    <tr>
                        <th>Company</th>
                        <th>Mentions</th>
                        <th>Last Seen</th>
                    </tr>
    """
    
    for company in analytics['top_companies'][:10]:  # Show top 10
        html_body += f"""
                    <tr>
                        <td>{company['name']}</td>
                        <td>{company['mention_count']}</td>
                        <td>{company['last_seen_date']}</td>
                    </tr>
        """
    
    html_body += """
                </table>
            </div>
            
            <div class="section">
                <h2>Top Locations</h2>
                <table>
                    <tr>
                        <th>Location</th>
                        <th>Mentions</th>
                        <th>Last Seen</th>
                    </tr>
    """
    
    for location in analytics['top_locations']:
        html_body += f"""
                    <tr>
                        <td>{location['name']}</td>
                        <td>{location['mention_count']}</td>
                        <td>{location['last_seen_date']}</td>
                    </tr>
        """
    
    html_body += """
                </table>
            </div>
            
            <div class="section">
                <h2>Project Types</h2>
                <table>
                    <tr>
                        <th>Type</th>
                        <th>Count</th>
                    </tr>
    """
    
    for ptype, count in analytics['project_types'].items():
        html_body += f"""
                    <tr>
                        <td>{ptype}</td>
                        <td>{count}</td>
                    </tr>
        """
    
    html_body += """
                </table>
            </div>
        </div>
    </body>
    </html>
    """
    
    # Attach HTML content
    msg.attach(MIMEText(html_body, 'html'))
    
    # Send email
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
        logging.info("Analytics report email sent successfully")
        return True
    except Exception as e:
        logging.error(f"Failed to send analytics email: {e}")
        return False

if __name__ == "__main__":
    main()