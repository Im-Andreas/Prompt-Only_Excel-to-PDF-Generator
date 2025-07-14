import pandas as pd
import matplotlib.pyplot as plt
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.fonts import addMapping
from PIL import Image
import os
import shutil

# Placeholder for reading Excel data
def read_excel(file_path):
    data = pd.read_excel(file_path)
    
    # Sort data alphabetically by the "Level 2" column (professor names)
    data = data.sort_values(by='Level 2', ascending=True)
    
    # Reset index after sorting
    data = data.reset_index(drop=True)
    
    # Create temp directory if it doesn't exist
    os.makedirs("temp", exist_ok=True)
    return data

# Function to create pie charts for each professor showing specialization distribution
def create_professor_pie_charts(data, specific_professor=None):
    # Find the column immediately before "Level 2" (specializations)
    level2_index = data.columns.get_loc('Level 2')
    specialization_col = data.columns[level2_index - 1]
    
    # Get unique professors
    all_professors = data['Level 2'].unique()
    
    # Filter professors based on input
    if specific_professor:
        if specific_professor in all_professors:
            professors = [specific_professor]
        else:
            return
    else:
        professors = all_professors
    
    # Create pie charts for selected professor(s)
    for professor in professors:
        if pd.isna(professor):  # Skip if professor name is NaN
            continue
            
        # Filter data for current professor
        prof_data = data[data['Level 2'] == professor]
        
        # Count specializations for this professor
        spec_counts = prof_data[specialization_col].value_counts()
        
        if len(spec_counts) > 0:  # Only create chart if there's data
            plt.figure(figsize=(12, 10))
            wedges, texts, autotexts = plt.pie(spec_counts.values, labels=spec_counts.index, autopct='%1.1f%%', startangle=90)
            
            # Increase font sizes for better readability
            for text in texts:
                text.set_fontsize(14)
                text.set_fontweight('bold')
            for autotext in autotexts:
                autotext.set_fontsize(12)
                autotext.set_fontweight('bold')
                autotext.set_color('white')
            
            plt.title('Student Specialization Distribution', 
                     fontsize=20, fontweight='bold', pad=30)
            plt.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle
            
            # Save chart with professor name (sanitize filename)
            safe_filename = "".join(c for c in professor if c.isalnum() or c in (' ', '-', '_')).rstrip()
            chart_filename = f"temp/pie_chart_{safe_filename}.png"
            plt.savefig(chart_filename, bbox_inches='tight', dpi=150)
            plt.close()
            
            # Generate individual PDF for this professor
            pdf_filename = f"output/report_{safe_filename}.pdf"
            generate_professor_pdf(pdf_filename, chart_filename, professor, spec_counts, len(prof_data), specialization_col, prof_data)
            
            # Clean up temporary files after PDF generation
            cleanup_temp_folder()
            

# Function to calculate proper image dimensions maintaining aspect ratio
def get_image_dimensions(image_path, max_width=400):
    """
    Calculate image dimensions maintaining aspect ratio
    """
    try:
        with Image.open(image_path) as img:
            original_width, original_height = img.size
            
        # Calculate aspect ratio
        aspect_ratio = original_height / original_width
        
        # Calculate dimensions maintaining aspect ratio
        width = max_width
        height = int(width * aspect_ratio)
        
        return width, height
    except:
        # Fallback dimensions if image can't be read
        return 400, 300
def parse_level3_data(level3_value):
    """
    Parse Level 3 data to extract course and year
    Example: "Matematică aplicată în economie-Curs-Anul 1" -> ("Matematică aplicată în economie-Curs", "Anul 1")
    """
    if pd.isna(level3_value):
        return None, None
    
    # Convert to string if not already
    level3_str = str(level3_value)
    
    # Split by '-' and look for year information
    parts = level3_str.split('-')
    
    # Find the part that contains "Anul"
    year = None
    course_parts = []
    
    for part in parts:
        part = part.strip()
        if "anul" in part.lower():
            year = part
        else:
            # Include all parts that are not year information (including Curs, Seminar, Laborator)
            course_parts.append(part)
    
    course = '-'.join(course_parts).strip() if course_parts else level3_str
    
    return course, year

# Function to generate detailed PDF for each professor
def generate_professor_pdf(output_path, chart_path, professor_name, spec_counts, total_students, specialization_col, prof_data):
    c = canvas.Canvas(output_path, pagesize=letter)
    
    # Register Arial Unicode MS or DejaVu fonts for UTF-8 support
    try:
        # Try to register a Unicode font (you may need to adjust the path)
        pdfmetrics.registerFont(TTFont('Unicode', 'arial.ttf'))
        unicode_font = 'Unicode'
    except:
        # Fallback to built-in fonts with limited Unicode support
        unicode_font = 'Helvetica'
    
    # NEW PAGE 1: TITLE PAGE WITH LOGO AND COMPLETION TRENDS
    # Add university logo
    logo_path = "assets/LOGO-ULBS_orizontal.png"
    if os.path.exists(logo_path):
        logo_width, logo_height = get_image_dimensions(logo_path, max_width=300)
        # Center the logo horizontally
        x_position = (letter[0] - logo_width) / 2
        c.drawImage(logo_path, x_position, 650, width=logo_width, height=logo_height)
    
    # Professor name (emphasized)
    c.setFont(unicode_font, 24)
    c.drawString(50, 600, "Performance Evaluation Report")
    
    c.setFont(unicode_font, 20)
    c.drawString(50, 570, f"Professor: {professor_name}")
    
    # Student completion statistics
    c.setFont(unicode_font, 14)
    c.drawString(50, 530, f"Total Students who Completed the Form: {total_students}")
    
    # Generate daily completion trend chart
    timestamp_col = 'Timestamp (dd/mm/yyyy)'
    if timestamp_col in data.columns:
        # Get timestamp data for this professor
        prof_timestamps = prof_data[timestamp_col].dropna()
        
        if len(prof_timestamps) > 0:
            # Convert timestamps to dates and count daily completions
            import matplotlib.dates as mdates
            from datetime import datetime, timedelta
            
            # Convert to datetime and extract dates
            dates = pd.to_datetime(prof_timestamps, errors='coerce').dt.date
            dates = dates.dropna()  # Remove any invalid dates
            
            if len(dates) > 0:
                daily_counts = dates.value_counts().sort_index()
                
                # Get the actual date range from the data
                min_date = daily_counts.index.min()
                max_date = daily_counts.index.max()
                
                # Use all available data without filtering
                period_data = daily_counts
                
                if len(period_data) > 0:
                    # Create line chart for daily completions
                    plt.figure(figsize=(14, 8))
                    
                    # Convert dates back to datetime for plotting
                    plot_dates = [datetime.combine(date, datetime.min.time()) for date in period_data.index]
                    
                    plt.plot(plot_dates, period_data.values, marker='o', linewidth=3, markersize=8, color='#2E86C1')
                    
                    plt.title('Daily Form Completion Trends', 
                             fontsize=18, fontweight='bold', pad=30)
                    plt.xlabel('Date', fontsize=16, fontweight='bold')
                    plt.ylabel('Number of Completions', fontsize=16, fontweight='bold')
                    
                    # Format x-axis with human-readable dates
                    plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%d %b'))
                    
                    # Adjust date locator based on data range
                    date_range_days = (max_date - min_date).days
                    if date_range_days <= 14:
                        plt.gca().xaxis.set_major_locator(mdates.DayLocator(interval=1))
                    elif date_range_days <= 30:
                        plt.gca().xaxis.set_major_locator(mdates.DayLocator(interval=2))
                    else:
                        plt.gca().xaxis.set_major_locator(mdates.WeekdayLocator(interval=1))
                    
                    plt.xticks(rotation=45, fontsize=14)
                    plt.yticks(fontsize=14)
                    
                    # Set y-axis to show only integer values
                    plt.gca().yaxis.set_major_locator(plt.MaxNLocator(integer=True))
                    
                    # Add grid
                    plt.grid(True, alpha=0.3)
                    
                    # Add value labels on points
                    for i, (date, count) in enumerate(zip(plot_dates, period_data.values)):
                        plt.annotate(f'{count}', (date, count), textcoords="offset points", 
                                   xytext=(0,12), ha='center', fontsize=12, fontweight='bold')
                    
                    plt.tight_layout()
                    
                    completion_chart_path = chart_path.replace('.png', '_completion_trends.png')
                    plt.savefig(completion_chart_path, bbox_inches='tight', dpi=150)
                    plt.close()
                    
                    # Add the completion trends chart to PDF
                    chart_width, chart_height = get_image_dimensions(completion_chart_path, max_width=500)
                    c.drawImage(completion_chart_path, 50, 200, width=chart_width, height=chart_height)
                    
                    # Add summary statistics with human-readable format
                    c.setFont(unicode_font, 10)
                    peak_date_formatted = period_data.idxmax().strftime('%d %B')
                    c.drawString(50, 180, f"Peak completion day: {peak_date_formatted} ({period_data.max()} completions)")
                    c.drawString(50, 165, f"Average daily completions: {int(round(period_data.mean()))}")
                    c.drawString(50, 150, f"Total days with responses: {len(period_data)}")
                    c.drawString(50, 135, f"Data range: {min_date.strftime('%d %B')} to {max_date.strftime('%d %B')}")
                else:
                    c.setFont(unicode_font, 12)
                    c.drawString(50, 400, "No completion data available")
                    c.drawString(50, 380, f"Available data range: {min_date.strftime('%d %B')} to {max_date.strftime('%d %B')}")
            else:
                c.setFont(unicode_font, 12)
                c.drawString(50, 400, "No valid timestamp data found for this professor")
        else:
            c.setFont(unicode_font, 12)
            c.drawString(50, 400, "No timestamp data available for this professor")
    else:
        c.setFont(unicode_font, 12)
        c.drawString(50, 400, f"Timestamp column '{timestamp_col}' not found in data")
    
    # Footer for title page
    c.setFont(unicode_font, 8)
    c.drawString(50, 50, f"Generated on: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # PAGE 2: SPECIALIZATION REPORT (formerly page 1)
    c.showPage()  # Start new page for specialization report
    
    # Title
    c.setFont(unicode_font, 26)
    c.drawString(50, 750, "Student Specializations")
    
    # Professor name
    c.setFont(unicode_font, 16)
    c.drawString(50, 720, f"Professor: {professor_name}")
    
    # Report details
    c.setFont(unicode_font, 12)
    c.drawString(50, 690, f"Total Students: {total_students}")
    c.drawString(50, 650, f"Number of Different Specializations: {len(spec_counts)}")
    
    # Specialization breakdown
    c.setFont(unicode_font, 14)
    c.drawString(50, 610, "Specialization Breakdown:")
    
    c.setFont(unicode_font, 10)
    y_position = 590
    for specialization, count in spec_counts.items():
        percentage = (count / total_students) * 100
        text = f"• {specialization}: {count} students ({percentage:.1f}%)"
        c.drawString(70, y_position, text)
        y_position -= 15
        if y_position < 320:  # Leave space for chart
            break
    
    # Add the main specialization chart (centered)
    chart_width, chart_height = get_image_dimensions(chart_path)
    x_position = (letter[0] - chart_width) / 2  # Center horizontally
    c.drawImage(chart_path, x_position, 50, width=chart_width, height=chart_height)
    
    # Parse Level 3 data for additional charts
    courses = []
    years = []
    
    # Calculate total professor responses (will be used across all pages)
    total_professor_responses = len(prof_data)
    
    # Find the attendance column (immediately after "Level 3")
    level3_index = data.columns.get_loc('Level 3')
    attendance_col = data.columns[level3_index + 1]
    workload_col = data.columns[level3_index + 2]
    
    # Teaching method columns - the 4 columns immediately after workload
    workload_index = data.columns.get_loc(workload_col)
    teaching_method_cols = []
    for i in range(1, 5):  # Get the next 4 columns after workload
        if workload_index + i < len(data.columns):
            teaching_method_cols.append(data.columns[workload_index + i])
    
    # Question columns - the 12 columns immediately after teaching methods
    question_cols = []
    for i in range(5, 17):  # Get the next 12 columns after workload (positions 5-16 after workload)
        if workload_index + i < len(data.columns):
            question_cols.append(data.columns[workload_index + i])
    
    # Comment columns - the 3 columns immediately after question columns
    comment_cols = []
    for i in range(17, 20):  # Get the next 3 columns after questions (positions 17-19 after workload)
        if workload_index + i < len(data.columns):
            comment_cols.append(data.columns[workload_index + i])
    
    for level3_value in prof_data['Level 3']:
        course, year = parse_level3_data(level3_value)
        if course:
            courses.append(course)
        if year:
            years.append(year)
    
    # PAGE 3: YEAR DISTRIBUTION
    if years:
        c.showPage()  # Start new page
        year_counts = pd.Series(years).value_counts()
        
        # Create year distribution chart
        plt.figure(figsize=(12, 10))
        wedges, texts, autotexts = plt.pie(year_counts.values, labels=year_counts.index, autopct='%1.1f%%', startangle=90)
        
        # Increase font sizes for better readability
        for text in texts:
            text.set_fontsize(14)
            text.set_fontweight('bold')
        for autotext in autotexts:
            autotext.set_fontsize(12)
            autotext.set_fontweight('bold')
            autotext.set_color('white')
        
        plt.title('Academic Year Distribution', 
                 fontsize=20, fontweight='bold', pad=30)
        plt.axis('equal')
        
        # Add subtitle with additional information
        years_no_response = total_professor_responses - len(years)
        
        year_chart_path = chart_path.replace('.png', '_years.png')
        plt.savefig(year_chart_path, bbox_inches='tight', dpi=150)
        plt.close()
        
        # Page 3 content
        c.setFont(unicode_font, 26)
        c.drawString(50, 750, "Academic Years")
        
        c.setFont(unicode_font, 16)
        c.drawString(50, 720, f"Professor: {professor_name}")
        
        c.setFont(unicode_font, 12)
        c.drawString(50, 690, f"Total Students for Professor: {total_professor_responses}")
        c.drawString(50, 670, f"Year Responses: {len(years)}")
        c.drawString(50, 650, f"Students with No Response: {years_no_response}")
        c.drawString(50, 630, f"Response Rate: {(len(years)/total_professor_responses)*100:.1f}%")
        
        # Year breakdown
        c.setFont(unicode_font, 14)
        c.drawString(50, 600, "Academic Year Breakdown:")
        
        c.setFont(unicode_font, 10)
        y_position = 580
        for year, count in year_counts.items():
            percentage_of_responses = (count / len(years)) * 100
            percentage_of_total = (count / total_professor_responses) * 100
            text = f"• {year}: {count} students ({percentage_of_responses:.1f}% of responses, {percentage_of_total:.1f}% of total)"
            c.drawString(70, y_position, text)
            y_position -= 15
            if y_position < 320:  # Leave space for chart
                break
        
        # Add no response information
        if years_no_response > 0:
            percentage_no_response = (years_no_response / total_professor_responses) * 100
            c.drawString(70, y_position, f"• No Response: {years_no_response} students ({percentage_no_response:.1f}% of total)")
        
        year_chart_width, year_chart_height = get_image_dimensions(year_chart_path)
        x_position = (letter[0] - year_chart_width) / 2  # Center horizontally
        c.drawImage(year_chart_path, x_position, 50, width=year_chart_width, height=year_chart_height)
    
    # PAGE 4: COURSE DISTRIBUTION
    if courses:
        c.showPage()  # Start new page
        course_counts = pd.Series(courses).value_counts()
        
        # Create course distribution chart
        plt.figure(figsize=(12, 10))
        wedges, texts, autotexts = plt.pie(course_counts.values, labels=course_counts.index, autopct='%1.1f%%', startangle=90)
        
        # Increase font sizes for better readability
        for text in texts:
            text.set_fontsize(14)
            text.set_fontweight('bold')
        for autotext in autotexts:
            autotext.set_fontsize(12)
            autotext.set_fontweight('bold')
            autotext.set_color('white')
        
        plt.title('Courses Distribution', 
                 fontsize=20, fontweight='bold', pad=30)
        plt.axis('equal')
        
        # Add subtitle with additional information
        courses_no_response = total_professor_responses - len(courses)
        
        course_chart_path = chart_path.replace('.png', '_courses.png')
        plt.savefig(course_chart_path, bbox_inches='tight', dpi=150)
        plt.close()
        
        # Page 4 content
        c.setFont(unicode_font, 26)
        c.drawString(50, 750, "Courses")
        
        c.setFont(unicode_font, 16)
        c.drawString(50, 720, f"Professor: {professor_name}")
        
        c.setFont(unicode_font, 12)
        c.drawString(50, 690, f"Total Students for Professor: {total_professor_responses}")
        c.drawString(50, 670, f"Course Responses: {len(courses)}")
        c.drawString(50, 650, f"Students with No Response: {courses_no_response}")
        c.drawString(50, 630, f"Response Rate: {(len(courses)/total_professor_responses)*100:.1f}%")
        
        # Course breakdown
        c.setFont(unicode_font, 14)
        c.drawString(50, 600, "Courses Breakdown:")
        
        c.setFont(unicode_font, 10)
        y_position = 580
        for course, count in course_counts.items():
            percentage_of_responses = (count / len(courses)) * 100
            percentage_of_total = (count / total_professor_responses) * 100
            text = f"• {course}: {count} evaluations ({percentage_of_responses:.1f}% of responses, {percentage_of_total:.1f}% of total)"
            c.drawString(70, y_position, text)
            y_position -= 15
            if y_position < 320:  # Leave space for chart
                break
        
        # Add no response information
        if courses_no_response > 0:
            percentage_no_response = (courses_no_response / total_professor_responses) * 100
            c.drawString(70, y_position, f"• No Response: {courses_no_response} students ({percentage_no_response:.1f}% of total)")
        
        course_chart_width, course_chart_height = get_image_dimensions(course_chart_path)
        x_position = (letter[0] - course_chart_width) / 2  # Center horizontally
        c.drawImage(course_chart_path, x_position, 50, width=course_chart_width, height=course_chart_height)
    
    # PAGE 5: ATTENDANCE DISTRIBUTION
    # Get attendance data for this professor
    attendance_data = prof_data[attendance_col].dropna()
    total_professor_responses = len(prof_data)
    attendance_no_response = total_professor_responses - len(attendance_data)
    
    if len(attendance_data) > 0:
        c.showPage()  # Start new page
        attendance_counts = attendance_data.value_counts().sort_index()
        
        # Create attendance distribution bar chart
        plt.figure(figsize=(12, 10))
        bars = plt.bar(range(len(attendance_counts)), attendance_counts.values, 
                      color='skyblue', edgecolor='navy', linewidth=1.5)
        
        # Customize the chart
        plt.title('Student Attendance Rate Distribution', 
                 fontsize=20, fontweight='bold', pad=30)
        plt.xlabel('Attendance Rate', fontsize=16, fontweight='bold')
        plt.ylabel('Number of Students', fontsize=16, fontweight='bold')
        
        # Set x-axis labels
        plt.xticks(range(len(attendance_counts)), attendance_counts.index, rotation=45, fontsize=14)
        plt.yticks(fontsize=14)
        
        # Add value labels on top of bars
        for i, bar in enumerate(bars):
            height = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                    f'{int(height)}', ha='center', va='bottom', fontweight='bold', fontsize=12)
        
        # Add grid for better readability
        plt.grid(axis='y', alpha=0.3, linestyle='--')
        
        plt.tight_layout()
        
        attendance_chart_path = chart_path.replace('.png', '_attendance.png')
        plt.savefig(attendance_chart_path, bbox_inches='tight', dpi=150)
        plt.close()
        
        # Page 5 content
        c.setFont(unicode_font, 26)
        c.drawString(50, 750, "Student Attendance")
        
        c.setFont(unicode_font, 16)
        c.drawString(50, 720, f"Professor: {professor_name}")
        
        c.setFont(unicode_font, 12)
        c.drawString(50, 690, f"Total Students for Professor: {total_professor_responses}")
        c.drawString(50, 670, f"Attendance Responses: {len(attendance_data)}")
        c.drawString(50, 650, f"Students with No Response: {attendance_no_response}")
        c.drawString(50, 630, f"Response Rate: {(len(attendance_data)/total_professor_responses)*100:.1f}%")
        
        # Attendance breakdown
        c.setFont(unicode_font, 14)
        c.drawString(50, 600, "Attendance Rate Breakdown:")
        
        c.setFont(unicode_font, 10)
        y_position = 580
        for attendance_rate, count in attendance_counts.items():
            percentage_of_responses = (count / len(attendance_data)) * 100
            percentage_of_total = (count / total_professor_responses) * 100
            text = f"• {attendance_rate}: {count} students ({percentage_of_responses:.1f}% of responses, {percentage_of_total:.1f}% of total)"
            c.drawString(70, y_position, text)
            y_position -= 15
            if y_position < 320:  # Leave space for chart
                break
        
        # Add no response information
        if attendance_no_response > 0:
            percentage_no_response = (attendance_no_response / total_professor_responses) * 100
            c.drawString(70, y_position, f"• No Response: {attendance_no_response} students ({percentage_no_response:.1f}% of total)")
        
        attendance_chart_width, attendance_chart_height = get_image_dimensions(attendance_chart_path)
        x_position = (letter[0] - attendance_chart_width) / 2  # Center horizontally
        c.drawImage(attendance_chart_path, x_position, 50, width=attendance_chart_width, height=attendance_chart_height)
    
    # PAGE 6: WORKLOAD DISTRIBUTION
    # Get workload data for this professor
    workload_data = prof_data[workload_col].dropna()
    workload_no_response = total_professor_responses - len(workload_data)
    
    if len(workload_data) > 0:
        c.showPage()  # Start new page
        
        # Define custom order for workload levels
        workload_order = ["Foarte mic", "Mic", "Mediu", "Mare", "Foarte mare"]
        
        # Get value counts and reorder according to custom order
        workload_counts_raw = workload_data.value_counts()
        workload_counts = pd.Series(dtype='int64')
        
        # Reorder according to the specified order, including zeros for missing categories
        for level in workload_order:
            if level in workload_counts_raw.index:
                workload_counts[level] = workload_counts_raw[level]
            else:
                workload_counts[level] = 0  # Add zero count for missing categories
        
        # Add any levels not in the predefined order at the end
        for level in workload_counts_raw.index:
            if level not in workload_order:
                workload_counts[level] = workload_counts_raw[level]
        
        # Create workload distribution bar chart
        plt.figure(figsize=(12, 10))
        bars = plt.bar(range(len(workload_counts)), workload_counts.values, 
                      color='lightcoral', edgecolor='darkred', linewidth=1.5)
        
        # Customize the chart
        plt.title('Student Workload Distribution', 
                 fontsize=20, fontweight='bold', pad=30)
        plt.xlabel('Workload Level', fontsize=16, fontweight='bold')
        plt.ylabel('Number of Students', fontsize=16, fontweight='bold')
        
        # Set x-axis labels
        plt.xticks(range(len(workload_counts)), workload_counts.index, rotation=45, fontsize=14)
        plt.yticks(fontsize=14)
        
        # Add value labels on top of bars
        for i, bar in enumerate(bars):
            height = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                    f'{int(height)}', ha='center', va='bottom', fontweight='bold', fontsize=12)
        
        # Add grid for better readability
        plt.grid(axis='y', alpha=0.3, linestyle='--')
        
        plt.tight_layout()
        
        workload_chart_path = chart_path.replace('.png', '_workload.png')
        plt.savefig(workload_chart_path, bbox_inches='tight', dpi=150)
        plt.close()
        
        # Page 6 content
        c.setFont(unicode_font, 26)
        c.drawString(50, 750, "Student Workload")
        
        c.setFont(unicode_font, 16)
        c.drawString(50, 720, f"Professor: {professor_name}")
        
        c.setFont(unicode_font, 12)
        c.drawString(50, 690, f"Total Students for Professor: {total_professor_responses}")
        c.drawString(50, 670, f"Workload Responses: {len(workload_data)}")
        c.drawString(50, 650, f"Students with No Response: {workload_no_response}")
        c.drawString(50, 630, f"Response Rate: {(len(workload_data)/total_professor_responses)*100:.1f}%")
        
        # Workload breakdown
        c.setFont(unicode_font, 14)
        c.drawString(50, 600, "Workload Level Breakdown:")
        
        c.setFont(unicode_font, 10)
        y_position = 580
        for workload_level, count in workload_counts.items():
            percentage_of_responses = (count / len(workload_data)) * 100
            percentage_of_total = (count / total_professor_responses) * 100
            text = f"• {workload_level}: {count} students ({percentage_of_responses:.1f}% of responses, {percentage_of_total:.1f}% of total)"
            c.drawString(70, y_position, text)
            y_position -= 15
            if y_position < 320:  # Leave space for chart
                break
        
        # Add no response information
        if workload_no_response > 0:
            percentage_no_response = (workload_no_response / total_professor_responses) * 100
            c.drawString(70, y_position, f"• No Response: {workload_no_response} students ({percentage_no_response:.1f}% of total)")
        
        workload_chart_width, workload_chart_height = get_image_dimensions(workload_chart_path)
        x_position = (letter[0] - workload_chart_width) / 2  # Center horizontally
        c.drawImage(workload_chart_path, x_position, 50, width=workload_chart_width, height=workload_chart_height)
    
    # PAGE 7: TEACHING METHODS DISTRIBUTION
    # Analyze teaching methods from the 4 columns after workload
    if teaching_method_cols and len(prof_data) > 0:
        c.showPage()  # Start new page
        
        total_responses = len(prof_data)
        
        # Count each teaching method type from the 4 columns
        method_counts = {
            'Predare CLASICĂ': 0,
            'Predare online SINCRONĂ': 0,
            'Predare online ASINCRONĂ': 0,
            'Predare MIXTĂ': 0
        }
        
        method_names_mapping = [
            'Predare CLASICĂ',
            'Predare online SINCRONĂ',
            'Predare online ASINCRONĂ',
            'Predare MIXTĂ'
        ]
        
        # Count non-null values in each teaching method column
        for i, col in enumerate(teaching_method_cols):
            if i < len(method_names_mapping):
                method_name = method_names_mapping[i]
                # Count non-null values (implemented methods)
                implemented_count = prof_data[col].notna().sum()
                method_counts[method_name] = implemented_count
        
        # Create teaching methods distribution bar chart
        method_names = list(method_counts.keys())
        method_values = list(method_counts.values())
        
        plt.figure(figsize=(14, 10))
        bars = plt.bar(range(len(method_names)), method_values, 
                      color=['#4CAF50', '#2196F3', '#FF9800', '#9C27B0'], 
                      edgecolor='black', linewidth=1.5, alpha=0.8)
        
        # Customize the chart
        plt.title('Teaching Methods Implementation', 
                 fontsize=20, fontweight='bold', pad=30)
        plt.xlabel('Teaching Method', fontsize=16, fontweight='bold')
        plt.ylabel('Number of Students Reporting Method', fontsize=16, fontweight='bold')
        
        # Set x-axis labels
        plt.xticks(range(len(method_names)), method_names, rotation=45, ha='right', fontsize=14)
        plt.yticks(fontsize=14)
        
        # Add value labels on top of bars with percentages
        for i, bar in enumerate(bars):
            height = bar.get_height()
            percentage = (height / total_responses) * 100 if total_responses > 0 else 0
            plt.text(bar.get_x() + bar.get_width()/2., height + 0.5,
                    f'{int(height)}\n({percentage:.1f}%)', 
                    ha='center', va='bottom', fontweight='bold', fontsize=12)
        
        # Add grid for better readability
        plt.grid(axis='y', alpha=0.3, linestyle='--')
        
        plt.tight_layout()
        
        teaching_chart_path = chart_path.replace('.png', '_teaching_methods.png')
        plt.savefig(teaching_chart_path, bbox_inches='tight', dpi=150)
        plt.close()
        
        # Page 7 content
        c.setFont(unicode_font, 26)
        c.drawString(50, 750, "Teaching Methods")
        
        c.setFont(unicode_font, 16)
        c.drawString(50, 720, f"Professor: {professor_name}")
        
        c.setFont(unicode_font, 12)
        c.drawString(50, 690, f"Total Student Responses: {total_responses}")
        c.drawString(50, 670, f"Number of Teaching Methods Analyzed: {len(teaching_method_cols)}")
        
        # Calculate overall statistics
        total_method_implementations = sum(method_values)
        
        c.drawString(50, 650, f"Total Method Implementations: {total_method_implementations}")
        
        # Teaching methods breakdown
        c.setFont(unicode_font, 14)
        c.drawString(50, 620, "Teaching Methods Breakdown:")
        
        c.setFont(unicode_font, 10)
        y_position = 600
        for i, (method_name, count) in enumerate(method_counts.items()):
            percentage = (count / total_responses) * 100 if total_responses > 0 else 0
            
            text = f"• {method_name}:"
            c.drawString(70, y_position, text)
            y_position -= 12
            c.drawString(90, y_position, f"Used by: {count} students ({percentage:.1f}%)")
            y_position -= 18
            
            if y_position < 320:  # Leave space for chart
                break
        
        teaching_chart_width, teaching_chart_height = get_image_dimensions(teaching_chart_path)
        x_position = (letter[0] - teaching_chart_width) / 2  # Center horizontally
        c.drawImage(teaching_chart_path, x_position, 50, width=teaching_chart_width, height=teaching_chart_height)
    
    # PAGES 8-19: INDIVIDUAL EVALUATION QUESTIONS ANALYSIS (PARETO CHARTS)
    # Create a separate page for each of the 12 evaluation questions
    if question_cols and len(prof_data) > 0:
        
        for q_index, col in enumerate(question_cols):
            c.showPage()  # Start new page for each question
            
            # Get the question text from the second row (index 1) of the original data
            if len(data) > 1:
                question_text = str(data.iloc[1, data.columns.get_loc(col)])
            else:
                question_text = f"Evaluation Question {q_index + 1}"
            
            # Get question data for this professor
            question_data = prof_data[col].dropna()
            total_responses_for_question = len(question_data)
            no_response_count = len(prof_data) - total_responses_for_question
            
            if total_responses_for_question > 0:
                # Convert to numeric and count grades 1-10
                numeric_data = pd.to_numeric(question_data, errors='coerce').dropna()
                
                if len(numeric_data) > 0:
                    # Count occurrences of each grade (1-10)
                    grade_counts = {}
                    for grade in range(1, 11):
                        count = (numeric_data == grade).sum()
                        if count > 0:  # Only include grades that have responses
                            grade_counts[grade] = count
                    
                    # Sort by count (descending for Pareto)
                    sorted_grades = sorted(grade_counts.items(), key=lambda x: x[1], reverse=True)
                    
                    if len(sorted_grades) > 0:  # Only proceed if there are grades with responses
                        # Prepare data for Pareto chart
                        grades = [str(item[0]) for item in sorted_grades]
                        counts = [item[1] for item in sorted_grades]
                        
                        # Calculate cumulative percentages
                        total_count = sum(counts)
                        cumulative_percentages = []
                        cumulative_sum = 0
                        for count in counts:
                            cumulative_sum += count
                            cumulative_percentages.append((cumulative_sum / total_count) * 100 if total_count > 0 else 0)
                        
                        # Create Pareto chart
                        fig, ax1 = plt.subplots(figsize=(14, 10))
                        
                        # Bar chart for grade counts
                        bars = ax1.bar(range(len(grades)), counts, color='lightblue', alpha=0.8, edgecolor='darkblue', linewidth=1.5)
                        ax1.set_xlabel('Grade (1-10)', fontsize=16, fontweight='bold')
                        ax1.set_ylabel('Number of Students', fontsize=16, fontweight='bold', color='darkblue')
                        ax1.tick_params(axis='y', labelcolor='darkblue', labelsize=14)
                        ax1.tick_params(axis='x', labelsize=14)
                        
                        # Set x-axis labels
                        ax1.set_xticks(range(len(grades)))
                        ax1.set_xticklabels(grades)
                        
                        # Add value labels on bars
                        for i, bar in enumerate(bars):
                            height = bar.get_height()
                            ax1.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                                    f'{int(height)}', ha='center', va='bottom', fontweight='bold', fontsize=12)
                        
                        # Line chart for cumulative percentage
                        ax2 = ax1.twinx()
                        line = ax2.plot(range(len(grades)), cumulative_percentages, color='red', marker='o', 
                                       linewidth=3, markersize=8, label='Cumulative %')
                        ax2.set_ylabel('Cumulative Percentage (%)', fontsize=16, fontweight='bold', color='red')
                        ax2.tick_params(axis='y', labelcolor='red', labelsize=14)
                        ax2.set_ylim(0, 100)
                        
                        # Add percentage labels on line points
                        for i, pct in enumerate(cumulative_percentages):
                            ax2.text(i, pct + 2, f'{pct:.1f}%', ha='center', va='bottom', 
                                    fontweight='bold', fontsize=12, color='red')
                        
                        # Add grid
                        ax1.grid(axis='y', alpha=0.3, linestyle='--')
                        
                        # Title
                        plt.title(f'Question {q_index + 1} - Grade Distribution (Pareto Analysis)', 
                                 fontsize=18, fontweight='bold', pad=30)
                        
                        # Calculate average score
                        avg_score = numeric_data.mean()
                        
                        plt.tight_layout()
                        
                        question_chart_path = chart_path.replace('.png', f'_question_{q_index + 1}.png')
                        plt.savefig(question_chart_path, bbox_inches='tight', dpi=150)
                        plt.close()
                        
                        # Page content
                        c.setFont(unicode_font, 24)
                        c.drawString(50, 750, f"Question {q_index + 1} - Grade Distribution Analysis")
                        
                        c.setFont(unicode_font, 16)
                        c.drawString(50, 720, f"Professor: {professor_name}")
                        
                        # Question text (truncated if too long)
                        c.setFont(unicode_font, 11)
                        if len(question_text) > 80:
                            question_text_display = question_text[:77] + "..."
                        else:
                            question_text_display = question_text
                        c.drawString(50, 690, f"Question: {question_text_display}")
                        
                        c.setFont(unicode_font, 12)
                        c.drawString(50, 660, f"Total Students for Professor: {len(prof_data)}")
                        c.drawString(50, 640, f"Students who Responded: {total_responses_for_question}")
                        c.drawString(50, 620, f"Students with No Response: {no_response_count}")
                        c.drawString(50, 600, f"Response Rate: {(total_responses_for_question/len(prof_data))*100:.1f}%")
                        c.drawString(50, 580, f"Average Score: {avg_score:.2f}/10")
                        
                        # Grade distribution breakdown
                        c.setFont(unicode_font, 14)
                        c.drawString(50, 550, "Grade Distribution (Pareto Order):")
                        
                        c.setFont(unicode_font, 10)
                        y_position = 530
                        for i, (grade, count) in enumerate(sorted_grades):  # All grades shown have responses > 0
                            percentage_of_responses = (count / total_responses_for_question) * 100
                            percentage_of_total = (count / len(prof_data)) * 100
                            cumulative_pct = cumulative_percentages[i]
                            
                            text = f"• Grade {grade}: {count} students ({percentage_of_responses:.1f}% of responses, {percentage_of_total:.1f}% of total) - Cumulative: {cumulative_pct:.1f}%"
                            c.drawString(70, y_position, text)
                            y_position -= 15
                            
                            if y_position < 340:  # Leave space for chart
                                break
                        
                        question_chart_width, question_chart_height = get_image_dimensions(question_chart_path)
                        x_position = (letter[0] - question_chart_width) / 2  # Center horizontally
                        c.drawImage(question_chart_path, x_position, 50, width=question_chart_width, height=question_chart_height)
                    
                    else:
                        # No grades with responses (should not happen if total_responses_for_question > 0)
                        c.setFont(unicode_font, 24)
                        c.drawString(50, 750, f"Question {q_index + 1} - No Valid Grades")
                        
                        c.setFont(unicode_font, 16)
                        c.drawString(50, 720, f"Professor: {professor_name}")
                        
                        c.setFont(unicode_font, 12)
                        c.drawString(50, 690, f"Question: {question_text}")
                        c.drawString(50, 660, f"No valid grades found for this question.")
                
                else:
                    # No valid numeric data for this question
                    c.setFont(unicode_font, 24)
                    c.drawString(50, 750, f"Question {q_index + 1} - No Valid Data")
                    
                    c.setFont(unicode_font, 16)
                    c.drawString(50, 720, f"Professor: {professor_name}")
                    
                    c.setFont(unicode_font, 12)
                    c.drawString(50, 690, f"Question: {question_text}")
                    c.drawString(50, 660, f"No valid numeric responses found for this question.")
            
            else:
                # No responses for this question
                c.setFont(unicode_font, 24)
                c.drawString(50, 750, f"Question {q_index + 1} - No Responses")
                
                c.setFont(unicode_font, 16)
                c.drawString(50, 720, f"Professor: {professor_name}")
                
                c.setFont(unicode_font, 12)
                c.drawString(50, 690, f"Question: {question_text}")
                c.drawString(50, 660, f"No responses found for this question.")
    
    # PAGES 20-22: COMMENTS ANALYSIS
    # Create pages for Pros, Cons, and "May Need Improvements" comments
    if comment_cols and len(prof_data) > 0:
        comment_section_names = [
            "Positive Aspects (Pros)",
            "Negative Aspects (Cons)", 
            "Areas of Improvement"
        ]
        
        for comment_index, col in enumerate(comment_cols):
            c.showPage()  # Start new page for each comment section
            
            section_name = comment_section_names[comment_index]
            
            # Get comment data for this professor
            comment_data = prof_data[col].dropna()
            # Remove empty strings and whitespace-only comments
            comment_data = comment_data[comment_data.str.strip() != '']
            
            total_comments = len(comment_data)
            no_comment_count = len(prof_data) - total_comments
            
            # Page header
            c.setFont(unicode_font, 26)
            c.drawString(50, 750, f"{section_name}")
            
            c.setFont(unicode_font, 16)
            c.drawString(50, 720, f"Professor: {professor_name}")
            
            c.setFont(unicode_font, 12)
            c.drawString(50, 690, f"Total Students for Professor: {len(prof_data)}")
            c.drawString(50, 670, f"Students with Comments: {total_comments}")
            c.drawString(50, 650, f"Students with No Comments: {no_comment_count}")
            c.drawString(50, 630, f"Comment Rate: {(total_comments/len(prof_data))*100:.1f}%")
            
            # Comments section header
            c.setFont(unicode_font, 14)
            c.drawString(50, 600, f"{section_name} - Student Comments:")
            
            # Display comments
            y_position = 570
            page_height_limit = 80  # Leave space for footer
            
            if total_comments > 0:
                for comment in comment_data.reset_index(drop=True):
                    # Calculate lines needed for this comment
                    comment_str = str(comment)
                    max_chars_per_line = 85  # Approximate characters per line
                    
                    # Split long comments into multiple lines
                    comment_lines = []
                    words = comment_str.split()
                    current_line = ""
                    
                    for word in words:
                        if len(current_line + " " + word) <= max_chars_per_line:
                            if current_line:
                                current_line += " " + word
                            else:
                                current_line = word
                        else:
                            if current_line:
                                comment_lines.append(current_line)
                            current_line = word
                    
                    if current_line:
                        comment_lines.append(current_line)
                    
                    # Check if we need a new page
                    lines_needed = len(comment_lines) + 2  # +2 for bullet and spacing
                    if y_position - (lines_needed * 12) < page_height_limit:
                        c.showPage()
                        
                        # Repeat header on new page
                        c.setFont(unicode_font, 26)
                        c.drawString(50, 750, f"{section_name} (continued)")
                        
                        c.setFont(unicode_font, 16)
                        c.drawString(50, 720, f"Professor: {professor_name}")
                        
                        y_position = 690
                    
                    # Draw bullet point with emphasis
                    c.setFont(unicode_font, 14)
                    c.setFillColorRGB(0.2, 0.4, 0.8)  # Blue color for bullet
                    c.drawString(50, y_position, "•")
                    c.setFillColorRGB(0, 0, 0)  # Reset to black
                    
                    # Draw comment text with proper indentation (aligned with bullet)
                    c.setFont(unicode_font, 10)
                    
                    for line_index, line in enumerate(comment_lines):
                        if line_index == 0:
                            # First line aligned horizontally with bullet
                            c.drawString(70, y_position, line)
                        else:
                            # Subsequent lines align with first line
                            c.drawString(70, y_position, line)
                        y_position -= 12
                    
                    # Add spacing between comments
                    y_position -= 15
            
            else:
                # No comments found
                c.setFont(unicode_font, 12)
                c.setFillColorRGB(0.5, 0.5, 0.5)  # Gray color for "no comments"
                c.drawString(70, y_position, "No comments provided by students for this section.")
                c.setFillColorRGB(0, 0, 0)  # Reset to black
            
            # Footer
            c.setFont(unicode_font, 8)
    
    c.save()

def cleanup_temp_folder():
    """
    Clean up the temp folder by removing all files inside it
    """
    temp_dir = "temp"
    if os.path.exists(temp_dir):
        # Remove all files and subdirectories in temp folder
        for filename in os.listdir(temp_dir):
            file_path = os.path.join(temp_dir, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)  # Remove file or link
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)  # Remove directory and its contents
            except Exception as e:
                print(f'Failed to delete {file_path}. Reason: {e}')

if __name__ == "__main__":
    # Create necessary directories
    os.makedirs("temp", exist_ok=True)
    os.makedirs("assets", exist_ok=True)
    os.makedirs("output", exist_ok=True)
    
    excel_file = "assets/QuestionPro-SR-RawData.xlsx"

    # Read data from Excel
    data = read_excel(excel_file)

    # Configuration: Choose to generate for all professors or a specific one
    # Set specific_professor to None for all professors, or provide a professor name
    #specific_professor = "ACU MUGUR"
    specific_professor = "OPREANA ALIN"  # Change this to a professor name if you want only one
    
    # Example: specific_professor = None  # Uncomment and modify to generate for specific professor
    
    # Create pie charts for selected professor(s) and generate individual PDFs
    create_professor_pie_charts(data, specific_professor)

    # Clean up temporary files
    cleanup_temp_folder()
