#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import time
import asyncio
import subprocess
from pathlib import Path
from pyppeteer import launch

# Chromium 다운로드 URL 수정 (최신 버전으로)
os.environ["PYPPETEER_CHROMIUM_REVISION"] = "1095492"  # 최신 안정 버전으로 조정

async def html_to_pdf(html_file, pdf_file=None):
    """
    HTML 파일을 PDF로 변환합니다 (Headless Chrome 사용).
    
    Args:
        html_file (str): 변환할 HTML 파일 경로 (index.xhtml)
        pdf_file (str, optional): 출력 PDF 파일 경로
    """
    html_path = Path(html_file)
    
    # 출력 파일 경로가 지정되지 않았다면 기본값 사용
    if pdf_file is None:
        # 기본 출력 경로는 HTML이 있는 디렉토리의 부모 디렉토리에 생성
        # ex: test2.html/index.xhtml -> test2.pdf
        parent_dir = html_path.parent
        if parent_dir.suffix == '.html':
            pdf_path = parent_dir.with_suffix('.pdf')
        else:
            pdf_path = html_path.with_suffix('.pdf')
    else:
        pdf_path = Path(pdf_file)
    
    # 입력 파일 존재 확인
    if not html_path.exists():
        print(f"오류: HTML 파일이 존재하지 않습니다: {html_path}")
        return
    
    # 파일이 생성된 직후에는 가끔 파일 잠금이 있을 수 있어 약간 대기
    print("파일 시스템 안정화를 위해 1초 대기...")
    await asyncio.sleep(1)
    
    print(f"PDF 변환 시작: {html_path} -> {pdf_path}")
    
    # 브라우저 변수를 밖에 선언하여 finally 블록에서 닫을 수 있도록 함
    browser = None
    
    try:
        # Chromium 시작 (executablePath 옵션으로 로컬 Chrome 사용 가능)
        browser_args = ['--no-sandbox', '--disable-setuid-sandbox']
        
        # 로컬 Chrome 경로 지정 (시스템에 따라 경로 조정 필요)
        chrome_paths = [
            # Windows 일반적인 경로
            "C:/Program Files/Google/Chrome/Application/chrome.exe",
            "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe",
            # MacOS 일반적인 경로
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
            # Linux 일반적인 경로
            "/usr/bin/google-chrome",
            "/usr/bin/chromium-browser",
        ]
        
        executable_path = None
        for path in chrome_paths:
            if os.path.exists(path):
                executable_path = path
                print(f"로컬 Chrome 발견: {path}")
                break
        
        launch_options = {
            'headless': True,
            'args': browser_args,
        }
        
        if executable_path:
            launch_options['executablePath'] = executable_path
            print("로컬 Chrome 사용")
        else:
            print("다운로드된 Chromium 사용 (없으면 다운로드 시도)")
        
        browser = await launch(**launch_options)
        page = await browser.newPage()
        
        # HTML 파일 로드 (file:/// 프로토콜 사용 - 슬래시 3개)
        # Windows 경로에서는 특수 처리 필요
        html_absolute_path = html_path.absolute()
        if sys.platform == 'win32':
            file_url = f"file:///{str(html_absolute_path).replace(os.sep, '/')}"
        else:
            file_url = f"file://{html_absolute_path}"
        
        print(f"HTML 파일 로드: {file_url}")
        # 페이지 로드 타임아웃 증가 및 대기 조건 조정
        await page.goto(file_url, {
            'waitUntil': 'networkidle0',
            'timeout': 60000  # 60초로 타임아웃 증가
        })
        
        # HTML 콘텐츠 확인
        page_content = await page.content()
        print(f"페이지 콘텐츠 길이: {len(page_content)} 바이트")
        
        # 빈 페이지인지 또는 디렉토리 목록인지 확인
        if len(page_content) < 200 or "디렉터리" in page_content or "Directory" in page_content:
            print("경고: 페이지가 비어있거나 디렉토리 목록입니다. HTML 코드를 직접 로드합니다.")
            
            try:
                # 파일 직접 읽기
                with open(html_path, 'r', encoding='utf-8') as f:
                    html_content = f.read()
                
                # CSS 파일 경로 확인 및 인라인화
                css_path = html_path.parent / "styles.css"
                css_content = ""
                if css_path.exists():
                    print(f"CSS 파일 발견: {css_path}")
                    with open(css_path, 'r', encoding='utf-8') as f:
                        css_content = f.read()
                
                # @page CSS 규칙 추가 - 여백 제거를 위한 핵심 수정 부분
                page_css = """
@page {
  margin: 0;
  padding: 0;
  size: A4;
}
html, body {
  margin: 0;
  padding: 0;
}
                """
                
                # HTML에 CSS 인라인 추가 (head 태그 내부에 style 태그 추가)
                if "<head>" in html_content:
                    if css_content:
                        html_with_css = html_content.replace(
                            "<head>", 
                            f"<head><style type=\"text/css\">{page_css}{css_content}</style>"
                        )
                    else:
                        html_with_css = html_content.replace(
                            "<head>", 
                            f"<head><style type=\"text/css\">{page_css}</style>"
                        )
                    html_content = html_with_css
                    print("@page CSS 규칙과 기존 CSS를 HTML에 인라인으로 추가했습니다.")
                
                # 직접 HTML 콘텐츠 설정
                await page.setContent(html_content, {
                    'waitUntil': 'networkidle0',
                    'timeout': 60000
                })
                print("HTML 콘텐츠를 직접 설정했습니다.")
                
                # 디버깅용 스크린샷 생성
                debug_screenshot = html_path.parent / "debug_screenshot.png"
                await page.screenshot({'path': str(debug_screenshot), 'fullPage': True})
                print(f"디버깅용 스크린샷 저장됨: {debug_screenshot}")
                
            except Exception as e:
                print(f"HTML 또는 CSS 파일 읽기 오류: {e}")
        else:
            # 페이지가 제대로 로드된 경우에도 @page CSS 규칙 주입
            await page.addStyleTag({
                'content': """
@page {
  margin: 0;
  padding: 0;
  border: none !important;
}
html, body {
  margin: 0;
  padding: 0;
}
                """
            })
            print("@page CSS 규칙을 페이지에 추가했습니다.")
        
        print("PDF 생성 중...")
        await page.pdf({
            'path': str(pdf_path),
            'format': 'A4',
            'printBackground': True,
            'margin': {'top': '0', 'right': '0', 'bottom': '0', 'left': '0'},
            'preferCSSPageSize': True  # CSS의 @page 규칙 사용 설정
        })
        
        if pdf_path.exists():
            print(f"PDF 변환 완료: {pdf_path}")
        else:
            print("PDF 변환 실패: 출력 파일을 찾을 수 없습니다.")
            
    except Exception as e:
        print(f"PDF 변환 중 오류 발생: {e}")
        import traceback
        traceback.print_exc()  # 상세 오류 정보 출력
    finally:
        # 브라우저 인스턴스가 생성되었는지 확인하고 닫기
        if browser:
            try:
                await browser.close()
            except:
                pass  # 브라우저 종료 오류는 무시

def convert_hwp_to_html(input_file, output_file=None):
    """
    HWP 파일을 HTML로 변환합니다.
    
    Args:
        input_file (str): 변환할 HWP 파일 경로
        output_file (str, optional): 출력 HTML 파일 경로
    
    Returns:
        Path: 생성된 HTML 파일 경로 또는 실패 시 None
    """
    input_path = Path(input_file)
    
    # 출력 파일 경로가 지정되지 않았다면 기본값 사용
    if output_file is None:
        output_dir = input_path.with_suffix('.html')
    else:
        output_dir = Path(output_file)
    
    # 입력 파일 존재 확인
    if not input_path.exists():
        print(f"오류: 입력 파일이 존재하지 않습니다: {input_path}")
        return None
    
    print(f"변환 시작: {input_path} -> {output_dir}")
    
    # 기존 HTML 디렉토리가 있다면 삭제 (hwp5html은 디렉토리를 생성함)
    if output_dir.exists():
        try:
            # 디렉토리인 경우 내용 삭제 (shutil 사용)
            import shutil
            if output_dir.is_dir():
                shutil.rmtree(output_dir)
                print(f"기존 디렉토리 삭제: {output_dir}")
            else:
                output_dir.unlink()
                print(f"기존 파일 삭제: {output_dir}")
        except Exception as e:
            print(f"기존 디렉토리/파일 삭제 실패: {e}")
    
    # PyHWP 명령행 도구를 사용하여 변환
    cmd = ["hwp5html", str(input_path), "--output", str(output_dir)]
    
    try:
        subprocess.run(cmd, check=True)
        
        # 파일이 생성되었는지 확인하기 전에 잠시 대기 (파일 시스템 안정화)
        import time
        time.sleep(1)
        
        # hwp5html은 디렉토리를 생성하고 그 안에 index.xhtml 파일을 생성함
        index_path = output_dir / "index.xhtml"
        
        if index_path.exists():
            print(f"변환 완료: {index_path} (in {output_dir})")
            return index_path
        else:
            print(f"변환 실패: index.xhtml 파일을 찾을 수 없습니다 (in {output_dir})")
            # 디렉토리가 존재하는지 확인
            if output_dir.exists() and output_dir.is_dir():
                print(f"디렉토리 내용 확인: {list(output_dir.iterdir())}")
            return None
            
    except subprocess.CalledProcessError as e:
        print(f"변환 중 오류 발생: {e}")
        return None
    except Exception as e:
        print(f"예외 발생: {e}")
        return None

def convert_hwp_to_pdf(input_file, output_pdf=None, output_html=None):
    """
    HWP 파일을 PDF로 변환합니다 (HTML을 중간 단계로 사용).
    
    Args:
        input_file (str): 변환할 HWP 파일 경로
        output_pdf (str, optional): 출력 PDF 파일 경로
        output_html (str, optional): 중간 HTML 파일 경로 (지정하지 않으면 임시 파일 사용)
    """
    input_path = Path(input_file)
    
    # 출력 PDF 파일 경로가 지정되지 않았다면 기본값 사용
    if output_pdf is None:
        pdf_path = input_path.with_suffix('.pdf')
    else:
        pdf_path = Path(output_pdf)
    
    # HWP를 HTML로 변환
    html_path = convert_hwp_to_html(input_file, output_html)
    
    if html_path:
        # HTML을 PDF로 변환 (비동기 함수 실행)
        asyncio.get_event_loop().run_until_complete(html_to_pdf(html_path, pdf_path))
    else:
        print("HWP에서 HTML 변환 실패로 PDF 생성을 건너뜁니다.")

async def main_async():
    """비동기 메인 함수"""
    # 명령행 인수 확인
    if len(sys.argv) < 2:
        print("사용법: python hwp2pdf.py [HWP파일경로] [PDF출력경로(선택사항)] [HTML출력경로(선택사항)]")
        return 1
    
    input_file = sys.argv[1]
    
    # 출력 파일 경로 (선택 사항)
    output_pdf = sys.argv[2] if len(sys.argv) > 2 else None
    output_html = sys.argv[3] if len(sys.argv) > 3 else None
    
    # HWP를 HTML로 변환
    html_path = convert_hwp_to_html(input_file, output_html)
    
    if html_path:
        # HTML을 PDF로 변환
        await html_to_pdf(html_path, output_pdf)
    else:
        print("HWP에서 HTML 변환 실패로 PDF 생성을 건너뜁니다.")
    
    return 0

def main():
    """메인 함수"""
    try:
        # 이벤트 루프 생성 및 실행
        if sys.platform == 'win32':
            asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
        
        sys.exit(asyncio.run(main_async()))
    except Exception as e:
        print(f"예기치 않은 오류: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()