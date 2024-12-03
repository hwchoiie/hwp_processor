import os
import re
import glob
import pyhwp
from datetime import datetime, timedelta

def extract_daily_info(file_path):
    """
    Extract daily content from .hwp file.
    """
    try:
        hwp = pyhwp.HWPDocument(file_path)
        content = hwp.get_body_text()
        # Customize extraction logic according to the format of your daily file
        return content
    except Exception as e:
        print(f"Failed to extract from {file_path}: {e}")
        return None

def compile_weekly_log(start_date, end_date, daily_files):
    """
    Compile weekly log from extracted daily information.
    """
    weekly_content = f"실시기간 {start_date.strftime('%Y년 %m월 %d일')} ~ {end_date.strftime('%Y년 %m월 %d일')}\n\n"
    for file in daily_files:
        day_content = extract_daily_info(file)
        if day_content:
            date_str = re.search(r'\d{4}-\d{2}-\d{2}', file).group()
            date = datetime.strptime(date_str, '%Y-%m-%d')
            day_name = ["월", "화", "수", "목", "금", "토", "일"][date.weekday()]
            weekly_content += f"{day_name}({date_str}) :\n{day_content}\n\n"
    return weekly_content

def write_to_hwp(file_path, content):
    """
    Write compiled content to .hwp file.
    """
    try:
        hwp = pyhwp.HWPDocument()
        hwp.add_body_text(content)
        hwp.save(file_path)
    except Exception as e:
        print(f"Failed to write to {file_path}: {e}")

def automate_weekly_log_generation(base_path, output_path, start_date, end_date):
    """
    Automate the process of generating weekly logs.
    """
    current_date = start_date
    while current_date < end_date:
        week_start = current_date
        week_end = current_date + timedelta(days=4)

        daily_files = sorted(glob.glob(os.path.join(base_path, f"*{week_start.strftime('%Y-%m-%d')}*.hwp")))
        weekly_log_content = compile_weekly_log(week_start, week_end, daily_files)

        output_file = os.path.join(output_path, f"주간일지_{week_start.strftime('%Y-%m-%d')}_~_{week_end.strftime('%Y-%m-%d')}.hwp")
        write_to_hwp(output_file, weekly_log_content)

        current_date += timedelta(days=7)

if __name__ == "__main__":
    base_path = "/path/to/daily/files"  # 경로 설정
    output_path = "/path/to/output/files"  # 출력 파일 경로 설정
    start_date = datetime(2024, 7, 8)  # 주간 일지 시작 날짜
    end_date = datetime(2024, 12, 31)  # 주간 일지 종료 날짜

    automate_weekly_log_generation(base_path, output_path, start_date, end_date)
