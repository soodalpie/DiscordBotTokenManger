import asyncio
import discord
from discord.ext import commands
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import json
import base64
import webbrowser
import os
from datetime import datetime
import tempfile
import html as html_lib
import re
import urllib.request
try:
    from tkinterweb import HtmlFrame as TkHtmlFrame
except Exception:
    TkHtmlFrame = None

TOKEN = None  # 로그인 폼에서 입력받음

intents = discord.Intents.default()
intents.guilds = True
intents.members = True
intents.message_content = True
bot = commands.Bot(command_prefix="!", intents=intents)
gui = None

def get_appdata_dir(app_name: str = '봇토큰관리') -> str:
    try:
        if os.name == 'nt':
            base = os.path.expanduser('~')
            base = os.environ.get('APPDATA', base)
            path = os.path.join(base, app_name)
        else:
            base = os.path.expanduser('~/.config')
            path = os.path.join(base, app_name)
        os.makedirs(path, exist_ok=True)
        return path
    except Exception:
        return os.getcwd()

@bot.event
async def on_ready():
    # GUI에 로그인 상태 전달 (메인 스레드에서 실행되도록 after 사용)
    if gui is not None:
        try:
            gui.root.after(0, gui.on_bot_ready)
        except Exception:
            pass

@bot.event
async def on_message(message: discord.Message):
    # 명령 처리 보존
    try:
        if gui is not None and message is not None:
            # 이제 봇 자신의 메시지도 기록하여 실시간 미리보기/트리에 반영
            try:
                rec = gui.message_to_rec(message)
                if message.guild is None:
                    gui.root.after(0, gui.ingest_dm_log, rec)
                else:
                    gui.root.after(0, gui.ingest_channel_log, rec)
            except Exception:
                pass
    finally:
        # 기존 명령 처리 유지
        await bot.process_commands(message)

def _start_bot(token: str):
    asyncio.run(bot.start(token))

class BotGUI:
    def __init__(self, bot):
        self.bot = bot
        self.root = tk.Tk()
        self.root.title("Discord Bot Manager")
        # 윈도우 크기/최소 크기 설정
        self.root.geometry("960x720")
        self.root.minsize(800, 600)
        self.ready_applied = False
        # 이미지 인라인(외부 → data URI) 사용 여부(기본 끔)
        self.inline_external_images = False
        # HTML 프레임 스크롤 상태(프레임 id → (y0,y1)) 및 하단 고정 여부
        self.html_last_yview = {}
        self.html_stick_bottom = {}
        self.html_frame_tag = {}  # 프레임 id → 'dm'|'ch'|'viewer'
        # 창 상태 저장 파일
        self.ui_state_path = os.path.join(get_appdata_dir('봇토큰관리'), 'ui_state.json')
        self.ui_state = self._ui_load_state()
        # 메인 창 위치 복원
        try:
            self._ui_restore_window(self.root, 'main_root', default_geo="960x720")
        except Exception:
            pass
        # 레이아웃 영역 구성
        self.header = ttk.Frame(self.root, padding=8)
        self.header.pack(fill='x')
        self.invite_area = ttk.Frame(self.root, padding=8)
        self.invite_area.pack(fill='x')
        self.search_area = ttk.Frame(self.root, padding=8)
        self.search_area.pack(fill='x')
        self.main_area = ttk.Frame(self.root, padding=8)
        self.main_area.pack(fill='both', expand=True)
        self.actions = ttk.Frame(self.root, padding=8)
        self.actions.pack(fill='x')
        # 메인 창 닫힘 시 위치 저장
        try:
            self.root.protocol("WM_DELETE_WINDOW", self._on_main_close)
        except Exception:
            pass
        # 길드 표시 문자열 → Guild 객체 매핑 및 리스트/디스플레이 보관
        self.guild_map = {}
        self.guild_list = []
        self.guild_displays = []
        self.selected_guild_id = None
        self._updating_guild_combo = False
        # 길드 선택 + 상태 + 새로고침 (헤더)
        self.guild_var = tk.StringVar(self.root)
        self.guild_combo = ttk.Combobox(self.header, textvariable=self.guild_var, width=40, state='readonly')
        self.guild_combo.pack(side='left')
        self.guild_combo.bind("<<ComboboxSelected>>", lambda e: self.on_guild_changed())
        self.status_var = tk.StringVar(self.root, value="로그인 대기 중...")
        self.status_label = ttk.Label(self.header, textvariable=self.status_var)
        self.status_label.pack(side='left', padx=10)
        self.refresh_btn = tk.Button(self.header, text="새로고침", command=self.refresh_all)
        self.refresh_btn.pack(side='right')
        # 뷰어(3번째 폼) 열기 버튼
        self.viewer_btn = tk.Button(self.header, text="뷰어", command=self.open_viewer)
        self.viewer_btn.pack(side='right', padx=6)
        # 초대 링크 표시/메뉴 (라벨 + 읽기전용 입력칸)
        self.invite_url = None
        ttk.Label(self.invite_area, text="초대 링크:").pack(side='left')
        self.invite_var = tk.StringVar(self.root, value="-")
        self.invite_entry = tk.Entry(self.invite_area, textvariable=self.invite_var, width=60, fg='blue')
        self.invite_entry.pack(side='left', fill='x', expand=True, padx=6)
        self.invite_entry.bind("<Double-Button-1>", lambda e: self.open_invite())
        self.invite_menu = tk.Menu(self.root, tearoff=0)
        self.invite_menu.add_command(label="링크 복사", command=lambda: self.copy_invite())
        self.invite_entry.bind("<Button-3>", lambda e: self.invite_menu.post(e.x_root, e.y_root))
        self.invite_buttons = ttk.Frame(self.invite_area)
        self.invite_buttons.pack(side='right')
        self.invite_refresh_btn = tk.Button(self.invite_buttons, text="초대 확인", command=self.load_invite)
        self.invite_refresh_btn.pack(side=tk.LEFT, padx=2)
        self.invite_create_btn = tk.Button(self.invite_buttons, text="초대 생성", command=self.create_invite)
        self.invite_create_btn.pack(side=tk.LEFT, padx=2)
        self.invite_copy_btn = tk.Button(self.invite_buttons, text="복사", command=self.copy_invite)
        self.invite_copy_btn.pack(side=tk.LEFT, padx=2)
        # 유저 검색
        self.search_var = tk.StringVar()
        ttk.Label(self.search_area, text="검색:").pack(side='left')
        self.search_entry = ttk.Entry(self.search_area, textvariable=self.search_var)
        self.search_entry.pack(side='left', fill='x', expand=True)
        self._search_after_id = None
        self.search_entry.bind("<KeyRelease>", self._on_search_key)
        # Enter 키로도 적용
        self.search_entry.bind("<Return>", self._on_search_key)
        try:
            self.search_var.trace_add('write', lambda *args: self._on_search_var_write())
        except Exception:
            pass
        # 검색 버튼
        ttk.Button(self.search_area, text="검색", command=self.apply_filter).pack(side='left', padx=6)
        # 멤버 캐시
        self.members_cache = []  # 전체 멤버 목록 (discord.Member)
        self.filtered_members = []  # 현재 필터 적용된 목록 (discord.Member)
        self.index_to_member = []  # 리스트박스 행 → member 매핑
        self.loading_guild_id = None  # 현재 로딩 중인 길드 ID(경쟁 상태 방지)
        # 로그인 전에는 로딩만 표시하고, 로그인 후 on_bot_ready에서 채움
        # 역할 만들기 버튼
        self.create_role_btn = tk.Button(self.actions, text="관리자 역할 만들기", command=self.create_admin_role)
        self.create_role_btn.pack(side='left', padx=4)
        # 유저 목록 + 스크롤
        self.user_scroll = ttk.Scrollbar(self.main_area, orient='vertical')
        self.user_listbox = tk.Listbox(self.main_area, yscrollcommand=self.user_scroll.set, selectmode=tk.EXTENDED)
        self.user_scroll.config(command=self.user_listbox.yview)
        self.user_listbox.pack(side='left', fill='both', expand=True)
        self.user_scroll.pack(side='left', fill='y')
        self.user_listbox.bind("<Button-3>", self.show_context_menu)
        # 컨텍스트 메뉴
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="관리자 권한 주기", command=self.give_admin_role)
        # 액션 버튼들
        self.kick_btn = tk.Button(self.actions, text="선택 유저 추방", command=self.kick_user)
        self.kick_btn.pack(side='left', padx=4)
        self.ban_btn = tk.Button(self.actions, text="선택 유저 차단", command=self.ban_user)
        self.ban_btn.pack(side='left', padx=4)
        # 선택/전체 제어 버튼
        self.select_all_btn = tk.Button(self.actions, text="전체 선택", command=self.select_all_users)
        self.select_all_btn.pack(side='left', padx=4)
        self.clear_sel_btn = tk.Button(self.actions, text="선택 해제", command=self.clear_selection)
        self.clear_sel_btn.pack(side='left', padx=4)
        # 관리자 권한 부여 버튼(선택/전체)
        self.give_admin_selected_btn = tk.Button(self.actions, text="선택 관리자권한", command=self.give_admin_role_selected)
        self.give_admin_selected_btn.pack(side='left', padx=4)
        self.give_admin_all_btn = tk.Button(self.actions, text="전체 관리자권한", command=self.give_admin_role_all)
        self.give_admin_all_btn.pack(side='left', padx=4)
        # 추방/차단 전체 버튼
        self.kick_all_btn = tk.Button(self.actions, text="전체 추방", command=self.kick_all_users)
        self.kick_all_btn.pack(side='left', padx=4)
        self.ban_all_btn = tk.Button(self.actions, text="전체 차단", command=self.ban_all_users)
        self.ban_all_btn.pack(side='left', padx=4)
        # 차단 목록 뷰어 버튼
        self.banlist_btn = tk.Button(self.actions, text="차단 목록", command=self.open_ban_viewer)
        self.banlist_btn.pack(side='left', padx=4)
        # DM 입력 영역
        self.dm_area = ttk.Frame(self.root, padding=8)
        self.dm_area.pack(fill='x')
        ttk.Label(self.dm_area, text="DM 메시지:").pack(side='left')
        self.dm_text = tk.Text(self.dm_area, height=3)
        self.dm_text.pack(side='left', fill='x', expand=True, padx=6)
        self.dm_sel_btn = tk.Button(self.dm_area, text="DM(선택)", command=self.send_dm_selected)
        self.dm_sel_btn.pack(side='left', padx=4)
        self.dm_all_btn = tk.Button(self.dm_area, text="DM(전체)", command=self.send_dm_all)
        self.dm_all_btn.pack(side='left', padx=4)
        # 로그 영역
        self.log_area = ttk.Frame(self.root, padding=8)
        self.log_area.pack(fill='both')
        self.log_scroll = ttk.Scrollbar(self.log_area, orient='vertical')
        self.log_text = tk.Text(self.log_area, height=8, yscrollcommand=self.log_scroll.set, state='disabled')
        self.log_scroll.config(command=self.log_text.yview)
        self.log_text.pack(side='left', fill='both', expand=True)
        self.log_scroll.pack(side='left', fill='y')
        # 컨트롤 기본 비활성화 (모든 위젯 생성 후 호출)
        self.set_controls_enabled(False)
        # 로그인 완료를 놓쳤을 수 있으므로 폴링 시작
        self.root.after(200, self.wait_for_ready)
        # 검색 폴링(키 이벤트 누락 대비)
        self._last_search_text = None
        try:
            self.root.after(300, self._search_poll)
        except Exception:
            pass

        # 메시지 로그 저장소/뷰어 핸들
        self.dm_logs = []
        self.channel_logs = []
        self.viewer_win = None
        self.dm_tree = None
        self.ch_tree = None
        self.viewer_html_holder = None  # 내장 HTML 뷰 프레임(옵션)
        # 채널 뷰 상태
        self.viewer_guild_var = tk.StringVar(self.root)
        self.viewer_channel_var = tk.StringVar(self.root)
        self.viewer_channel_tree = None
        self.viewer_selected_guild_id = None
        self.viewer_selected_channel_id = None
        self.viewer_channel_messages = []  # dict 리스트 [{time, author, content, id, created_at}, ...]
        self.viewer_channel_oldest = None  # datetime 또는 message
        # 뷰어 길드 콤보 상태/매핑
        self._viewer_updating_guild_combo = False
        self.viewer_guild_map = {}
        self.viewer_guild_list = []
        self.viewer_guild_displays = []
        # 실시간 HTML 프레임 참조
        self.dm_html_frame = None
        self.ch_html_frame = None
        self.viewer_html_frame = None
        # HTML 임베드 미리보기 기본 비활성화(트리 표시 우선)
        self.embed_preview = False
        # 외부 HTML 미리보기 창(종류별 보관: 'dm'|'ch'|'viewer' → (win, frame))
        self.live_html = {'dm': None, 'ch': None, 'viewer': None}
        # DM 전송 UI 상태
        self.dm_guild_var = tk.StringVar(self.root)
        self.dm_user_var = tk.StringVar(self.root)
        self.dm_msg_var = tk.StringVar(self.root)
        self.dm_guild_combo = None
        self.dm_user_combo = None
        self.dm_members_list = []
        self.dm_members_displays = []
        # DM 히스토리: 채널별 가장 오래된 메시지 포인터
        self.dm_history_oldest_by_channel = {}
        # 트리 자동 스크롤 상태
        self.dm_autoscroll = True
        self.ch_autoscroll = True
        self.viewer_autoscroll = True

    # ===== 메시지 로그 수집/표시 =====
    def ingest_dm_log(self, rec: dict):
        try:
            self.dm_logs.append(rec)
            if self.dm_tree is not None and self._is_viewer_alive():
                iid = self._append_tree_row(self.dm_tree, (rec.get('time',''), rec.get('author',''), rec.get('content','')))
                if self.dm_autoscroll and iid:
                    try:
                        self.dm_tree.see(iid)
                    except Exception:
                        pass
            # HTML 미리보기 갱신
            self._update_live_html('dm')
        except Exception:
            pass

    def ingest_channel_log(self, rec: dict):
        try:
            self.channel_logs.append(rec)
            if self.ch_tree is not None and self._is_viewer_alive():
                iid = self._append_tree_row(self.ch_tree, (
                    rec.get('time',''),
                    f"{rec.get('guild','')}#{rec.get('channel','')}",
                    rec.get('author',''),
                    rec.get('content',''),
                ))
                if self.ch_autoscroll and iid:
                    try:
                        self.ch_tree.see(iid)
                    except Exception:
                        pass
            # 선택 채널 실시간 반영
            try:
                if (self.viewer_channel_tree is not None
                    and self._is_viewer_alive()
                    and rec.get('channel_id') == self.viewer_selected_channel_id):
                    iid2 = self._append_tree_row(self.viewer_channel_tree, (
                         rec.get('time',''), rec.get('author',''), rec.get('content','')
                     ))
                    self.viewer_channel_messages.append(rec)
                    # oldest 업데이트
                    self._viewer_update_oldest_from_list()
                    if self.viewer_autoscroll and iid2:
                        try:
                            self.viewer_channel_tree.see(iid2)
                        except Exception:
                            pass
                    # 채널 뷰 HTML 갱신
                    self._update_live_html('viewer')
            except Exception:
                pass
            # 채널 전체 HTML 갱신
            self._update_live_html('ch')
        except Exception:
            pass

    def _is_viewer_alive(self):
        try:
            return self.viewer_win is not None and bool(self.viewer_win.winfo_exists())
        except Exception:
            return False

    def open_viewer(self):
        if self._is_viewer_alive():
            try:
                self.viewer_win.lift()
                self.viewer_win.focus_force()
            except Exception:
                pass
            return
        self._build_viewer()

    def _build_viewer(self):
        self.viewer_win = tk.Toplevel(self.root)
        self.viewer_win.title("메시지 뷰어")
        # 위치 복원
        try:
            self._ui_restore_window(self.viewer_win, 'viewer_win', default_geo="960x600")
        except Exception:
            self.viewer_win.geometry("960x600")

        # 상단 도구영역
        tools = ttk.Frame(self.viewer_win, padding=6)
        tools.pack(fill='x')
        ttk.Button(tools, text="DM HTML 보기", command=lambda: self._show_html('dm')).pack(side='left', padx=3)
        ttk.Button(tools, text="채널 HTML 보기", command=lambda: self._show_html('ch')).pack(side='left', padx=3)

        # 탭
        nb = ttk.Notebook(self.viewer_win)
        nb.pack(fill='both', expand=True)

        dm_tab = ttk.Frame(nb)
        ch_tab = ttk.Frame(nb)
        viewer_tab = ttk.Frame(nb)
        nb.add(dm_tab, text="DM")
        nb.add(ch_tab, text="채널")
        nb.add(viewer_tab, text="채널 뷰")

        # DM 컨트롤 + 트리
        dm_top = ttk.Frame(dm_tab, padding=6)
        dm_top.pack(fill='x')
        ttk.Label(dm_top, text="길드:").pack(side='left')
        self.dm_guild_combo = ttk.Combobox(dm_top, textvariable=self.dm_guild_var, width=30, state='readonly')
        self.dm_guild_combo.pack(side='left', padx=4)
        self.dm_guild_combo.bind("<<ComboboxSelected>>", lambda e: self._dm_on_guild_changed())
        ttk.Button(dm_top, text="멤버 불러오기", command=self._dm_on_guild_changed).pack(side='left', padx=4)
        ttk.Label(dm_top, text="대상:").pack(side='left')
        self.dm_user_combo = ttk.Combobox(dm_top, textvariable=self.dm_user_var, width=30, state='normal')
        self.dm_user_combo.pack(side='left', padx=4)
        try:
            self.dm_user_combo.bind('<KeyRelease>', self._dm_on_user_type)
        except Exception:
            pass
        # DM 히스토리 로딩 버튼
        ttk.Button(dm_top, text="최근 불러오기", command=self._dm_load_latest).pack(side='left', padx=4)
        ttk.Button(dm_top, text="이전 100개", command=self._dm_load_older).pack(side='left', padx=4)
        dm_send = ttk.Frame(dm_tab, padding=6)
        dm_send.pack(fill='x')
        ttk.Label(dm_send, text="DM 메시지:").pack(side='left')
        self.dm_entry = ttk.Entry(dm_send, textvariable=self.dm_msg_var)
        self.dm_entry.pack(side='left', fill='x', expand=True, padx=6)
        ttk.Button(dm_send, text="보내기", command=self._dm_send).pack(side='left', padx=4)
        try:
            self.dm_entry.bind('<Return>', lambda e: self._dm_send())
        except Exception:
            pass
        # DM 트리
        dm_cols = ("시간", "사용자", "내용")
        self.dm_tree = ttk.Treeview(dm_tab, columns=dm_cols, show='headings')
        for c, w in [("시간", 140), ("사용자", 180), ("내용", 600)]:
            self.dm_tree.heading(c, text=c)
            self.dm_tree.column(c, width=w, anchor='w')
        self.dm_tree.pack(fill='both', expand=True)
        try:
            self.dm_tree.configure(yscrollcommand=lambda a,b: self._on_tree_yview('dm', a, b))
            # 마우스휠/키 스크롤 시 상태 갱신(보조)
            for ev in ('<MouseWheel>', '<Button-4>', '<Button-5>', '<KeyPress-Up>', '<KeyPress-Down>', '<Prior>', '<Next>'):
                self.dm_tree.bind(ev, lambda e: self._update_autoscroll_from_view('dm'))
        except Exception:
            pass
        for r in self.dm_logs:
            self._append_tree_row(self.dm_tree, (r.get('time',''), r.get('author',''), r.get('content','')))
        # 초기에는 하단 고정
        self._scroll_tree_to_bottom(self.dm_tree)
        # 우클릭 컨텍스트 메뉴(선택 삭제)
        try:
            self.dm_tree_menu = tk.Menu(dm_tab, tearoff=0)
            self.dm_tree_menu.add_command(label="선택 메시지 삭제", command=self._dm_delete_selected)
            self.dm_tree.bind('<Button-3>', self._on_dm_tree_context)
        except Exception:
            pass
        # DM 실시간 HTML 미리보기(옵션)
        if TkHtmlFrame is not None and getattr(self, 'embed_preview', False):
            self.dm_html_frame = TkHtmlFrame(dm_tab, messages_enabled=False)
            self.dm_html_frame.pack(fill='both', expand=True)
            self._html_mark_frame(self.dm_html_frame, 'dm')
            self._html_bind_external_links(self.dm_html_frame)
            self._html_bind_scroll_state(self.dm_html_frame)
            self._html_set_content(self.dm_html_frame, self._build_html('dm'))
        # DM 길드 콤보 초기화(메인 선택과 동기화)
        self._dm_load_guilds()

        # 채널 트리
        ch_cols = ("시간", "채널", "작성자", "내용")
        self.ch_tree = ttk.Treeview(ch_tab, columns=ch_cols, show='headings')
        for c, w in [("시간", 140), ("채널", 220), ("작성자", 160), ("내용", 520)]:
            self.ch_tree.heading(c, text=c)
            self.ch_tree.column(c, width=w, anchor='w')
        self.ch_tree.pack(fill='both', expand=True)
        try:
            self.ch_tree.configure(yscrollcommand=lambda a,b: self._on_tree_yview('ch', a, b))
            for ev in ('<MouseWheel>', '<Button-4>', '<Button-5>', '<KeyPress-Up>', '<KeyPress-Down>', '<Prior>', '<Next>'):
                self.ch_tree.bind(ev, lambda e: self._update_autoscroll_from_view('ch'))
        except Exception:
            pass
        for r in self.channel_logs:
            chan = f"{r.get('guild','')}#{r.get('channel','')}"
            self._append_tree_row(self.ch_tree, (r.get('time',''), chan, r.get('author',''), r.get('content','')))
        self._scroll_tree_to_bottom(self.ch_tree)
        # 채널 실시간 HTML 미리보기(옵션)
        if TkHtmlFrame is not None and getattr(self, 'embed_preview', False):
            self.ch_html_frame = TkHtmlFrame(ch_tab, messages_enabled=False)
            self.ch_html_frame.pack(fill='both', expand=True)
            self._html_mark_frame(self.ch_html_frame, 'ch')
            self._html_bind_external_links(self.ch_html_frame)
            self._html_bind_scroll_state(self.ch_html_frame)
            self._html_set_content(self.ch_html_frame, self._build_html('ch'))

        # 채널 뷰(길드/채널 선택 → 메시지 로드)
        top = ttk.Frame(viewer_tab, padding=6)
        top.pack(fill='x')
        ttk.Label(top, text="길드:").pack(side='left')
        self.viewer_guild_combo = ttk.Combobox(top, textvariable=self.viewer_guild_var, width=35, state='readonly')
        self.viewer_guild_combo.pack(side='left', padx=4)
        self.viewer_guild_combo.bind("<<ComboboxSelected>>", lambda e: self._viewer_on_guild_changed())
        ttk.Label(top, text="채널:").pack(side='left')
        self.viewer_channel_combo = ttk.Combobox(top, textvariable=self.viewer_channel_var, width=35, state='readonly')
        self.viewer_channel_combo.pack(side='left', padx=4)
        ttk.Button(top, text="최신 불러오기", command=self._viewer_load_latest).pack(side='left', padx=4)
        ttk.Button(top, text="이전 100개", command=self._viewer_load_older).pack(side='left', padx=4)
        ttk.Button(top, text="HTML 보기", command=lambda: self._show_html('viewer')).pack(side='left', padx=4)
        # 채널 전송 UI
        send_bar = ttk.Frame(viewer_tab, padding=6)
        send_bar.pack(fill='x')
        self.viewer_send_var = tk.StringVar(self.root)
        ttk.Label(send_bar, text="메시지:").pack(side='left')
        self.viewer_send_entry = ttk.Entry(send_bar, textvariable=self.viewer_send_var)
        self.viewer_send_entry.pack(side='left', fill='x', expand=True, padx=6)
        ttk.Button(send_bar, text="전송", command=self._viewer_send_message).pack(side='left', padx=4)
        try:
            # 엔터 시 전송(체크박스 연동)
            self.viewer_enter_on_var = tk.BooleanVar(self.root, value=False)
            self.viewer_enter_cb = tk.Checkbutton(send_bar, text="엔터 시 전송", variable=self.viewer_enter_on_var)
            self.viewer_enter_cb.pack(side='left', padx=6)
            self.viewer_send_entry.bind('<Return>', self._on_viewer_enter)
        except Exception:
            pass


        cols = ("시간", "작성자", "내용")
        self.viewer_channel_tree = ttk.Treeview(viewer_tab, columns=cols, show='headings')
        for c, w in [("시간", 140), ("작성자", 180), ("내용", 580)]:
            self.viewer_channel_tree.heading(c, text=c)
            self.viewer_channel_tree.column(c, width=w, anchor='w')
        self.viewer_channel_tree.pack(fill='both', expand=True)
        try:
            self.viewer_channel_tree.configure(yscrollcommand=lambda a,b: self._on_tree_yview('viewer', a, b))
            for ev in ('<MouseWheel>', '<Button-4>', '<Button-5>', '<KeyPress-Up>', '<KeyPress-Down>', '<Prior>', '<Next>'):
                self.viewer_channel_tree.bind(ev, lambda e: self._update_autoscroll_from_view('viewer'))
        except Exception:
            pass
        # 우클릭 컨텍스트 메뉴(선택 삭제)
        try:
            self.viewer_tree_menu = tk.Menu(viewer_tab, tearoff=0)
            self.viewer_tree_menu.add_command(label="선택 메시지 삭제", command=self._viewer_delete_selected_menu)
            self.viewer_channel_tree.bind('<Button-3>', self._on_viewer_tree_context)
        except Exception:
            pass
        # 채널 뷰 실시간 HTML 미리보기(옵션)
        if TkHtmlFrame is not None and getattr(self, 'embed_preview', False):
            self.viewer_html_frame = TkHtmlFrame(viewer_tab, messages_enabled=False)
            self.viewer_html_frame.pack(fill='both', expand=True)
            self._html_mark_frame(self.viewer_html_frame, 'viewer')
            self._html_bind_external_links(self.viewer_html_frame)
            self._html_bind_scroll_state(self.viewer_html_frame)
            self._html_set_content(self.viewer_html_frame, self._build_html('viewer'))

        # 길드 콤보 초기화(메인 선택과 동기화)
        self._viewer_load_guilds(initial=True)

        # 창 닫힘 처리
        def _on_close_inner():
            try:
                # 위치 저장
                self._ui_save_window(self.viewer_win, 'viewer_win')
                self.dm_tree = None
                self.ch_tree = None
            finally:
                try:
                    self.viewer_win.destroy()
                except Exception:
                    pass
                self.viewer_win = None
        self.viewer_win.protocol("WM_DELETE_WINDOW", _on_close_inner)

    def _append_tree_row(self, tree: ttk.Treeview, row):
        try:
            return tree.insert('', 'end', values=row)
        except Exception:
            return None

    def _scroll_tree_to_bottom(self, tree: ttk.Treeview):
        try:
            children = tree.get_children()
            if children:
                tree.see(children[-1])
        except Exception:
            pass

    def _on_tree_yview(self, kind: str, first, last):
        # 스크롤 위치에 따라 자동 하단 고정 토글
        try:
            f = float(first); l = float(last)
            at_bottom = (l >= 0.999)
            if kind == 'dm':
                self.dm_autoscroll = at_bottom
            elif kind == 'ch':
                self.ch_autoscroll = at_bottom
            elif kind == 'viewer':
                self.viewer_autoscroll = at_bottom
        except Exception:
            pass

    def _update_autoscroll_from_view(self, kind: str):
        # 이벤트 기반 보조 업데이트
        try:
            tree = {'dm': self.dm_tree, 'ch': self.ch_tree, 'viewer': self.viewer_channel_tree}.get(kind)
            if tree is None:
                return
            f, l = tree.yview()
            self._on_tree_yview(kind, f, l)
        except Exception:
            pass

    # 메시지 객체를 통합 레코드(dict)로 변환
    def message_to_rec(self, m: discord.Message) -> dict:
        try:
            ts = m.created_at
            ts_str = ts.strftime('%Y-%m-%d %H:%M:%S') if isinstance(ts, datetime) else str(ts)
        except Exception:
            ts_str = ''
        author = str(getattr(m, 'author', '') or '')
        try:
            avatar = getattr(getattr(m, 'author', None), 'display_avatar', None)
            avatar_url = ''
            if avatar:
                try:
                    # discord.py v2: with_format/with_size
                    avatar_url = str(avatar.with_format('png').with_size(64).url)
                except Exception:
                    try:
                        # 일부 버전: replace(format=..., size=...)
                        avatar_url = str(avatar.replace(format='png', size=64).url)
                    except Exception:
                        avatar_url = str(getattr(avatar, 'url', '') or '')
                # 최종 폴백: webp → png 문자열 치환
                if 'webp' in (avatar_url or ''):
                    avatar_url = re.sub(r"\.webp(\?.*)?$", r".png\1", avatar_url)
                    avatar_url = avatar_url.replace('format=webp', 'format=png')
        except Exception:
            avatar_url = ''
        content = getattr(m, 'content', '') or ''
        # 텍스트를 HTML로 가공(멘션/이모지/코드블럭)
        try:
            content_text_html = self._render_text_html_from_message(m)
        except Exception:
            content_text_html = html_lib.escape(content).replace('\n', '<br>')
        # 첨부
        atts = []
        try:
            for a in getattr(m, 'attachments', []) or []:
                atts.append({
                    'url': getattr(a, 'url', ''),
                    'filename': getattr(a, 'filename', ''),
                    'content_type': getattr(a, 'content_type', ''),
                    'size': getattr(a, 'size', None),
                })
        except Exception:
            pass
        # 임베드(간략)
        embeds = []
        try:
            for e in getattr(m, 'embeds', []) or []:
                embeds.append({
                    'title': getattr(e, 'title', ''),
                    'description': getattr(e, 'description', ''),
                    'url': getattr(e, 'url', ''),
                })
        except Exception:
            pass
        # 스티커(간략)
        stickers = []
        try:
            for s in getattr(m, 'stickers', []) or []:
                stickers.append({'name': getattr(s, 'name', '')})
        except Exception:
            pass
        # 답장(참조)
        ref_info = None
        try:
            ref = getattr(m, 'reference', None)
            if ref and getattr(ref, 'resolved', None):
                rm = ref.resolved
                ref_author = str(getattr(rm, 'author', '') or '')
                ref_snip = (getattr(rm, 'content', '') or '')[:150]
                ref_info = {'author': ref_author, 'snippet': ref_snip}
        except Exception:
            ref_info = None
        # 리액션
        reactions = []
        try:
            for r in getattr(m, 'reactions', []) or []:
                emoji = str(getattr(r, 'emoji', ''))
                count = getattr(r, 'count', 0)
                reactions.append({'emoji': emoji, 'count': count})
        except Exception:
            pass

        rec = {
            'id': getattr(m, 'id', None),
            'time': ts_str,
            'author': author,
            'author_id': getattr(getattr(m, 'author', None), 'id', ''),
            'author_avatar': avatar_url,
            'content': content,
            'content_text_html': content_text_html,
            'attachments': atts,
            'embeds': embeds,
            'stickers': stickers,
            'reference': ref_info,
            'reactions': reactions,
        }
        if getattr(m, 'guild', None) is None:
            rec.update({
                'scope': 'dm',
                'channel': str(getattr(m, 'channel', '')),
                'channel_id': getattr(getattr(m, 'channel', None), 'id', ''),
            })
        else:
            rec.update({
                'scope': 'guild',
                'guild': m.guild.name,
                'guild_id': m.guild.id,
                'channel': getattr(m.channel, 'name', str(m.channel)),
                'channel_id': getattr(m.channel, 'id', ''),
            })
        # HTML 컨텐츠 생성
        rec['content_html'] = self._render_rich_html_for_rec(rec)
        return rec

    def _render_rich_html_for_rec(self, rec: dict) -> str:
        def esc(x):
            try:
                return html_lib.escape(str(x) if x is not None else '')
            except Exception:
                return ''
        # 본문(텍스트 → HTML 변환본 우선 사용)
        body = rec.get('content_text_html')
        if not body:
            body = esc(rec.get('content', '')).replace('\n', '<br>')
        # 첨부(이미지는 인라인, 그 외 링크)
        for a in rec.get('attachments', []) or []:
            url = a.get('url', '')
            fn = a.get('filename', '')
            ct = (a.get('content_type') or '').lower()
            # 파일명/URL 기준으로 확장자 판별(쿼리스트립)
            lower_fn = (fn or '').lower()
            lower_url = (url or '').lower().split('?', 1)[0]
            url_ext_img = lower_url.endswith(('.png', '.jpg', '.jpeg', '.gif', '.webp'))
            fn_ext_img = lower_fn.endswith(('.png', '.jpg', '.jpeg', '.gif', '.webp'))
            # 이미지 여부: Content-Type 또는 확장자 둘 중 하나라도 이미지면 True
            is_img = (ct.startswith('image/') or url_ext_img or fn_ext_img)
            # webp 여부는 확장자 기준으로만 판정(헤더 오검출 방지)
            is_webp = (lower_url.endswith('.webp') or lower_fn.endswith('.webp'))
            if is_img and url and not is_webp:
                # 이미지 + 클릭 시 외부 브라우저 열기 + 텍스트 링크(폴백)
                body += (
                    f"<div class='att'>"
                    f"<a href='{esc(url)}' target='_blank' rel='noopener'>"
                    f"<img src='{esc(url)}' alt='{esc(fn)}' width='680' style='display:block; max-width:100%; height:auto; border-radius:6px; margin-top:6px;'>"
                    f"</a>"
                    f"<div class='att-link'><a href='{esc(url)}' target='_blank' rel='noopener'>{esc(fn or '열기')}</a></div>"
                    f"</div>"
                )
            else:
                # webp 또는 비이미지는 링크로 표시(외부 브라우저에서 열기)
                label = fn or url
                body += f"<div class='att'><a href='{esc(url)}' target='_blank' rel='noopener'>{esc(label)}</a></div>"
        # 임베드(간략 카드)
        for e in rec.get('embeds', []) or []:
            title = esc(e.get('title', ''))
            desc = esc(e.get('description', ''))
            url = esc(e.get('url', ''))
            if url:
                title_html = f"<a href='{url}' target='_blank'>{title or url}</a>"
            else:
                title_html = title
            card = f"<div class='embed'><div class='et'>{title_html}</div>"
            if desc:
                card += f"<div class='ed'>{desc}</div>"
            card += "</div>"
            body += card
        # 스티커
        for s in rec.get('stickers', []) or []:
            body += f"<div class='sticker'>:{esc(s.get('name',''))}:</div>"
        # 답장(참조)
        ref = rec.get('reference')
        if ref:
            ra = esc(ref.get('author',''))
            rs = esc(ref.get('snippet',''))
            body = f"<div class='reply'><span class='r-author'>{ra}</span> {rs}</div>" + body
        # 리액션 바
        reacts = rec.get('reactions') or []
        if reacts:
            chips = []
            for r in reacts:
                chips.append(f"<span class='react'>{esc(r.get('emoji',''))} {esc(r.get('count',''))}</span>")
            body += f"<div class='reactions'>{''.join(chips)}</div>"
        # 아바타/헤더 포함한 메시지 블록
        avatar = esc(rec.get('author_avatar') or '')
        avatar_img = f"<img class='avatar' src='{avatar}'>" if avatar else "<div class='avatar avatar-ph'></div>"
        author = esc(rec.get('author',''))
        uid = esc(rec.get('author_id',''))
        time_str = esc(rec.get('time',''))
        header = f"<div class='hdr'><span class='author'>{author}</span> <span class='uid'>({uid})</span><span class='time'>{time_str}</span></div>"
        return f"<div class='msg'>{avatar_img}<div class='body'>{header}<div class='content'>{body}</div></div></div>"

    def _render_text_html_from_message(self, m: discord.Message) -> str:
        # 멘션/커스텀 이모지/코드블럭/인라인코드 처리 후 HTML 반환
        text = getattr(m, 'content', '') or ''
        # 멘션 대체
        try:
            for u in getattr(m, 'mentions', []) or []:
                for tok in (f"<@{u.id}>", f"<@!{u.id}>"):
                    text = text.replace(tok, f"@{u.display_name}")
            for r in getattr(m, 'role_mentions', []) or []:
                text = text.replace(f"<@&{r.id}>", f"@{r.name}")
            for ch in getattr(m, 'channel_mentions', []) or []:
                text = text.replace(f"<#{ch.id}>", f"#{ch.name}")
        except Exception:
            pass
        # 커스텀 이모지 → 이미지
        def _repl_emoji(mo):
            anim = mo.group(1)
            name = mo.group(2)
            eid = mo.group(3)
            # tkinterweb 호환을 위해 정적 이모지는 png 사용
            ext = 'gif' if (anim and anim.lower() == 'a') else 'png'
            url = f"https://cdn.discordapp.com/emojis/{eid}.{ext}?size=44&quality=lossless"
            return f"<img class='emoji' src='{html_lib.escape(url)}' alt=':{html_lib.escape(name)}:'>"
        try:
            text = re.sub(r"<(a?):([A-Za-z0-9_~]+):(\d+)>", _repl_emoji, text)
        except Exception:
            pass
        # 코드블럭/인라인코드 처리
        return self._convert_markdown_basic(text)

    def _convert_markdown_basic(self, text: str) -> str:
        # ```lang\n ... ``` 블럭 → <pre><code>
        blocks = []
        def _repl_block(mo):
            lang = (mo.group(1) or '').strip()
            code = mo.group(2)
            idx = len(blocks)
            blocks.append((lang, code))
            return f"[[[CODEBLOCK{idx}]]]"
        try:
            text2 = re.sub(r"```([^\n`]*)\n([\s\S]*?)```", _repl_block, text)
        except Exception:
            text2 = text
        # escape 후 줄바꿈 처리
        html_txt = html_lib.escape(text2)
        # 인라인 코드
        try:
            html_txt = re.sub(r"`([^`]+)`", lambda m: f"<code>{html_lib.escape(m.group(1))}</code>", html_txt)
        except Exception:
            pass
        html_txt = html_txt.replace('\n', '<br>')
        # 블럭 복원
        for i, (lang, code) in enumerate(blocks):
            code_html = f"<pre><code class='lang-{html_lib.escape(lang)}'>{html_lib.escape(code)}</code></pre>"
            html_txt = html_txt.replace(f"[[[CODEBLOCK{i}]]]", code_html)
        return html_txt

    def _build_html(self, which: str) -> str:
        # which: 'dm' | 'ch' | 'viewer'
        if which == 'dm':
            recs = self.dm_logs
            title = 'DM 로그'
        elif which == 'ch':
            recs = self.channel_logs
            title = '채널 로그'
        else:
            recs = self.viewer_channel_messages
            title = '채널 뷰'
        msgs = []
        for r in recs:
            try:
                # 항상 최신 템플릿으로 다시 렌더링(첨부/스타일 반영)
                html_snippet = self._render_rich_html_for_rec(r)
            except Exception:
                # 폴백 최소 렌더링
                html_snippet = self._render_rich_html_for_rec({
                    'author': r.get('author',''),
                    'author_avatar': '',
                    'time': r.get('time',''),
                    'content': r.get('content',''),
                    'attachments': [], 'embeds': [], 'stickers': []
                })
            msgs.append(html_snippet)
        html = f"""
        <html><head><meta charset='utf-8'>
        <style>
        body {{ background:#2b2d31; color:#dbdee1; margin:0; padding:16px; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Noto Sans KR', Arial, 'Apple SD Gothic Neo', '맑은 고딕', sans-serif; }}
        /* TkinterWeb 호환: flex 대신 float 레이아웃 사용 */
        .msg {{ margin:14px 0; }}
        .avatar {{ width:40px; height:40px; border-radius:50%; background:#1e1f22; float:left; margin-right:12px; }}
        .avatar-ph {{ width:40px; height:40px; border-radius:50%; background:#3f4147; float:left; margin-right:12px; }}
        .body {{ display:block; overflow:visible; }}
        .hdr {{ margin-bottom:2px; }}
        .author {{ font-weight:600; color:#fff; }}
        .uid {{ color:#949ba4; font-size:12px; }}
        .time {{ color:#949ba4; font-size:12px; }}
        .content {{ margin-top:4px; line-height:1.4; clear: both; }}
        .content code {{ background:#1e1f22; border:1px solid #3f4147; padding:1px 4px; border-radius:4px; font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, 'Liberation Mono', monospace; font-size:12px; }}
        pre {{ background:#1e1f22; border:1px solid #3f4147; padding:8px; border-radius:6px; overflow:auto; }}
        .content a {{ color:#00a8fc; text-decoration:none; }}
        .content a:hover {{ text-decoration:underline; }}
        .att img {{ max-width:100%; height:auto; border-radius:6px; margin-top:6px; display:block; }}
        .att a {{ color:#00a8fc; }}
        .att-link {{ margin-top:4px; font-size:12px; }}
        .embed {{ border-left:4px solid #5865f2; background:#313338; padding:8px; border-radius:4px; margin-top:6px; }}
        .embed .et {{ font-weight:600; margin-bottom:4px; }}
        .embed .ed {{ white-space:pre-wrap; color:#dbdee1; }}
        .sticker {{ display:inline-block; background:#383a40; padding:4px 8px; border-radius:6px; margin-top:6px; }}
        .reply {{ border-left:3px solid #4e5058; color:#b5bac1; padding-left:8px; margin-bottom:6px; }}
        .reply .r-author {{ color:#fff; font-weight:600; margin-right:6px; }}
        .reactions {{ margin-top:6px; display:flex; gap:6px; flex-wrap:wrap; }}
        .react {{ background:#2e3035; border:1px solid #3f4147; border-radius:12px; padding:2px 8px; color:#dbdee1; font-size:12px; }}
        .emoji {{ width:22px; height:22px; vertical-align:-5px; }}
        h3 {{ margin:0 0 12px 0; color:#fff; }}
        </style></head>
        <body>
        <h3>{title}</h3>
        {''.join(msgs)}
        </body></html>
        """
        return html

    def _show_html(self, which: str):
        # tkinterweb이 있으면 재사용 가능한 미리보기 창을 유지하고 실시간 갱신된다.
        try:
            try:
                from tkinterweb import HtmlFrame as TkHtmlFrame
            except Exception:
                TkHtmlFrame = None
            html = self._build_html(which)
            # 필요 시에만 이미지 인라인 처리
            if getattr(self, 'inline_external_images', False):
                html = self._inline_external_images(html)
            if TkHtmlFrame is not None:
                # 이미 열려 있으면 업데이트하고 앞으로 가져오기
                entry = self.live_html.get(which)
                if entry:
                    win, fr = entry
                    try:
                        if win is not None and bool(win.winfo_exists()):
                            self._html_set_content(fr, html)
                            try:
                                win.lift(); win.focus_force()
                            except Exception:
                                pass
                            return
                    except Exception:
                        pass
                    # 무효 핸들은 정리
                    self.live_html[which] = None
                # 새 창 생성
                parent = self.viewer_win if self.viewer_win and self.viewer_win.winfo_exists() else self.root
                win = tk.Toplevel(parent)
                title_map = {'dm': 'DM HTML', 'ch': '채널 HTML', 'viewer': '채널 뷰 HTML'}
                win.title(title_map.get(which, 'HTML 미리보기'))
                # 위치 복원
                try:
                    self._ui_restore_window(win, f'html_{which}', default_geo="900x600")
                except Exception:
                    win.geometry("900x600")
                fr = TkHtmlFrame(win, messages_enabled=False)
                fr.pack(fill='both', expand=True)
                self._html_mark_frame(fr, which)
                self._html_bind_external_links(fr)
                self._html_bind_scroll_state(fr)
                self._html_set_content(fr, html)
                # 닫힘 시 핸들 정리
                def _on_close():
                    try:
                        self._ui_save_window(win, f'html_{which}')
                        self.live_html[which] = None
                    finally:
                        try:
                            win.destroy()
                        except Exception:
                            pass
                win.protocol("WM_DELETE_WINDOW", _on_close)
                self.live_html[which] = (win, fr)
            else:
                # 임시파일로 저장 후 브라우저로 오픈(이 경우 실시간 갱신 불가)
                with tempfile.NamedTemporaryFile('w', delete=False, suffix='.html', encoding='utf-8') as f:
                    f.write(html)
                    path = f.name
                webbrowser.open(f"file://{path}")
        except Exception as e:
            messagebox.showerror("오류", f"HTML 보기 중 오류: {e}")

    # ===== 채널 뷰 보조 메서드 =====
    def _viewer_on_guild_changed(self):
        # 길드 선택 시 채널 목록 갱신
        if self._viewer_updating_guild_combo:
            return
        try:
            # 콤보 인덱스 기준으로 안정적으로 선택
            idx = self.viewer_guild_combo.current()
            g = self.viewer_guild_list[idx] if (isinstance(idx, int) and 0 <= idx < len(self.viewer_guild_list)) else None
            self.viewer_selected_guild_id = g.id if g else None
            # 먼저 현재 뷰 초기화(이후 자동 로드 시 덮어쓰지 않도록)
            self._viewer_clear_channel_view()
            chans = []
            if g:
                for ch in getattr(g, 'text_channels', []):
                    try:
                        perms = ch.permissions_for(g.me)
                        if getattr(perms, 'view_channel', False) and getattr(perms, 'read_message_history', False):
                            chans.append((ch.name, ch.id))
                    except Exception:
                        continue
            chan_disps = [f"{name} ({cid})" for name, cid in chans]
            self.viewer_channel_combo['values'] = chan_disps
            if chan_disps:
                self.viewer_channel_combo.current(0)
                self.viewer_channel_var.set(chan_disps[0])
                # 자동으로 첫 채널 최신 메시지 로드
                self._viewer_load_latest()
        except Exception:
            pass

    def _viewer_load_guilds(self, initial=False):
        # 메인 선택과 동기화하여 길드 콤보 채우기
        self.viewer_guild_list = list(self.bot.guilds)
        self.viewer_guild_displays = [f"{g.name} ({g.id})" for g in self.viewer_guild_list]
        self.viewer_guild_map = {d: g for d, g in zip(self.viewer_guild_displays, self.viewer_guild_list)}
        # 타겟 index 결정: 메인 선택 → 이전 뷰어 선택 → 0
        target_idx = 0 if self.viewer_guild_list else -1
        # 메인 콤보 기준 현재 선택 길드 ID 확인
        cur_guild = None
        try:
            cur_guild = self.get_selected_guild()
        except Exception:
            cur_guild = None
        sel_id = getattr(cur_guild, 'id', None)
        if sel_id is not None:
            for i, g in enumerate(self.viewer_guild_list):
                if g.id == sel_id:
                    target_idx = i
                    break
        elif self.viewer_selected_guild_id is not None:
            for i, g in enumerate(self.viewer_guild_list):
                if g.id == self.viewer_selected_guild_id:
                    target_idx = i
                    break
        self._viewer_updating_guild_combo = True
        try:
            self.viewer_guild_combo['values'] = self.viewer_guild_displays
            if target_idx >= 0 and self.viewer_guild_list:
                self.viewer_guild_combo.current(target_idx)
                self.viewer_guild_var.set(self.viewer_guild_displays[target_idx])
        finally:
            self._viewer_updating_guild_combo = False
        # 초기 길드 채널 목록 로드
        if initial:
            self._viewer_on_guild_changed()

    def _dm_load_guilds(self):
        try:
            glist = list(self.bot.guilds)
            displays = [f"{g.name} ({g.id})" for g in glist]
            self.dm_guild_combo['values'] = displays
            # 메인 선택 길드로 기본 선택
            cur = None
            try:
                cur = self.get_selected_guild()
            except Exception:
                cur = None
            target_idx = 0 if glist else -1
            if cur is not None:
                for i, g in enumerate(glist):
                    if g.id == cur.id:
                        target_idx = i
                        break
            if target_idx >= 0 and glist:
                self.dm_guild_combo.current(target_idx)
                self.dm_guild_var.set(displays[target_idx])
                # 자동으로 멤버 목록 로드
                self._dm_on_guild_changed()
        except Exception:
            pass

    def _viewer_clear_channel_view(self):
        try:
            if self.viewer_channel_tree is not None:
                for it in self.viewer_channel_tree.get_children():
                    self.viewer_channel_tree.delete(it)
            self.viewer_channel_messages = []
            self.viewer_selected_channel_id = None
            self.viewer_channel_oldest = None
        except Exception:
            pass

    def _viewer_parse_selected_channel(self):
        disp = (self.viewer_channel_var.get() or '').strip()
        if disp.endswith(')') and '(' in disp:
            try:
                cid = int(disp.rsplit('(', 1)[1].rstrip(')'))
                return cid
            except Exception:
                return None
        return None

    def _viewer_load_latest(self):
        gid = self.viewer_selected_guild_id
        cid = self._viewer_parse_selected_channel()
        if not gid or not cid:
            messagebox.showerror("오류", "길드/채널을 선택하세요.")
            return
        self._viewer_clear_channel_view()
        self.viewer_selected_channel_id = cid
        g = next((x for x in self.bot.guilds if x.id == gid), None)
        ch = None
        if g:
            ch = next((c for c in getattr(g, 'text_channels', []) if c.id == cid), None)
        if not ch:
            messagebox.showerror("오류", "채널을 찾을 수 없습니다.")
            return
        self.log(f"채널 뷰: 최신 불러오기 시작 → {g.name}#{ch.name}")
        fut = asyncio.run_coroutine_threadsafe(self._fetch_channel_history(ch, limit=100, before=None), self.bot.loop)
        def _done(f):
            try:
                msgs = f.result()
            except Exception:
                msgs = []
            self.root.after(0, self._viewer_apply_messages, msgs, True)
        fut.add_done_callback(_done)

    def _viewer_load_older(self):
        gid = self.viewer_selected_guild_id
        cid = self.viewer_selected_channel_id
        if not gid or not cid:
            messagebox.showerror("오류", "먼저 최신 메시지를 불러오세요.")
            return
        g = next((x for x in self.bot.guilds if x.id == gid), None)
        ch = None
        if g:
            ch = next((c for c in getattr(g, 'text_channels', []) if c.id == cid), None)
        if not ch:
            messagebox.showerror("오류", "채널을 찾을 수 없습니다.")
            return
        before = None
        try:
            # 가장 오래된 메시지 시간/객체 사용
            if self.viewer_channel_oldest and hasattr(self.viewer_channel_oldest, 'created_at'):
                before = self.viewer_channel_oldest
        except Exception:
            before = None
        self.log("채널 뷰: 이전 100개 불러오기")
        fut = asyncio.run_coroutine_threadsafe(self._fetch_channel_history(ch, limit=100, before=before), self.bot.loop)
        def _done(f):
            try:
                msgs = f.result()
            except Exception:
                msgs = []
            self.root.after(0, self._viewer_apply_messages, msgs, False)
        fut.add_done_callback(_done)

    async def _fetch_channel_history(self, channel, limit=100, before=None):
        out = []
        try:
            it = channel.history(limit=limit, before=before)
            async for m in it:
                out.append(m)
        except Exception:
            pass
        # 최신→과거로 오므로 뒤집어 시간순 정렬
        out.reverse()
        return out

    def _viewer_apply_messages(self, messages, replace: bool):
        if replace:
            self._viewer_clear_channel_view()
            self.viewer_selected_channel_id = self._viewer_parse_selected_channel()
        rows = []
        for m in messages:
            try:
                rec = self.message_to_rec(m)
            except Exception:
                # 폴백 최소 정보
                rec = self._channel_rec(m)
                rec['content_html'] = self._render_rich_html_for_rec({
                    'author': rec.get('author',''),
                    'author_avatar': '',
                    'time': rec.get('time',''),
                    'content': rec.get('content',''),
                    'attachments': [], 'embeds': [], 'stickers': []
                })
            rows.append(rec)
        # 트리에 추가
        try:
            last_iid = None
            for r in rows:
                last_iid = self._append_tree_row(self.viewer_channel_tree, (r['time'], r['author'], r['content']))
            # 최신 불러오기(replace=True) 또는 자동 하단 고정 상태면 하단 유지
            if (replace or self.viewer_autoscroll) and last_iid:
                try:
                    self.viewer_channel_tree.see(last_iid)
                except Exception:
                    pass
        except Exception:
            pass
        self.viewer_channel_messages.extend(rows)
        self._viewer_update_oldest_from_list(messages)
        # 채널 뷰 HTML 갱신(내장/외부 모두)
        self._update_live_html('viewer')
        # 작성자 콤보 갱신
        self._viewer_refresh_authors_combo()

    def _viewer_rebuild_tree_from_messages(self):
        try:
            if self.viewer_channel_tree is None:
                return
            for it in self.viewer_channel_tree.get_children():
                self.viewer_channel_tree.delete(it)
            last_iid = None
            for r in self.viewer_channel_messages:
                last_iid = self._append_tree_row(self.viewer_channel_tree, (
                    r.get('time',''), r.get('author',''), r.get('content','')
                ))
            if self.viewer_autoscroll and last_iid:
                try:
                    self.viewer_channel_tree.see(last_iid)
                except Exception:
                    pass
            # HTML/작성자 콤보 갱신
            self._update_live_html('viewer')
            self._viewer_refresh_authors_combo()
        except Exception:
            pass

    # ===== HTML 헬퍼 =====
    def _html_set_content(self, frame, html: str):
        # 스크롤 상태 보존 후 콘텐츠 갱신
        key = id(frame)
        y0 = None
        try:
            # 현재 스크롤 기억
            yv = frame.html.yview() if hasattr(frame, 'html') else frame.yview()
            if isinstance(yv, (list, tuple)) and len(yv) >= 2:
                y0 = float(yv[0])
                self.html_last_yview[key] = (float(yv[0]), float(yv[1]))
        except Exception:
            pass
        try:
            frame.set_content(html)
        except Exception:
            try:
                frame.load_html(html)
            except Exception:
                pass
        # 레이아웃 완료 후 스크롤 복원(하단 고정 또는 위치 유지)
        def _restore():
            try:
                stick = bool(self.html_stick_bottom.get(key, False))
                sv = self.html_last_yview.get(key)
                if hasattr(frame, 'html'):
                    if stick:
                        frame.html.yview_moveto(1.0)
                    elif sv and isinstance(sv, (list, tuple)):
                        frame.html.yview_moveto(max(0.0, min(1.0, float(sv[0]))))
                else:
                    if stick:
                        frame.yview_moveto(1.0)
                    elif sv and isinstance(sv, (list, tuple)):
                        frame.yview_moveto(max(0.0, min(1.0, float(sv[0]))))
            except Exception:
                pass
        try:
            frame.after(50, _restore)
        except Exception:
            _restore()

    # ===== UI 상태 저장/복원 =====
    def _ui_load_state(self) -> dict:
        try:
            with open(self.ui_state_path, 'r', encoding='utf-8') as f:
                return json.load(f) or {}
        except Exception:
            return {}

    def _ui_save_state(self):
        try:
            with open(self.ui_state_path, 'w', encoding='utf-8') as f:
                json.dump(self.ui_state or {}, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _ui_save_window(self, win, key: str):
        try:
            win.update_idletasks()
            state = None
            try:
                state = win.state()
            except Exception:
                state = 'normal'
            geo = win.winfo_geometry()
            self.ui_state[key] = {'geo': geo, 'state': state}
            self._ui_save_state()
        except Exception:
            pass

    def _ui_restore_window(self, win, key: str, default_geo: str = "800x600"):
        try:
            info = (self.ui_state or {}).get(key) or {}
            geo = info.get('geo') or default_geo
            try:
                win.geometry(geo)
            except Exception:
                win.geometry(default_geo)
            st = info.get('state')
            if st in ('zoomed', 'normal'):
                try:
                    win.state(st)
                except Exception:
                    pass
        except Exception:
            win.geometry(default_geo)

    def _on_main_close(self):
        try:
            self._ui_save_window(self.root, 'main_root')
        except Exception:
            pass
        try:
            self.root.destroy()
        except Exception:
            pass

    def _html_bind_external_links(self, frame):
        # HtmlFrame 링크 클릭 → 외부 브라우저로 열기
        try:
            frame.config(on_link_click=lambda url: webbrowser.open(url))
        except Exception:
            try:
                frame.on_link_click(lambda url: webbrowser.open(url))
            except Exception:
                pass

    # ===== HTML 스크롤 상태 유틸 =====
    def _html_mark_frame(self, frame, tag: str):
        try:
            self.html_frame_tag[id(frame)] = tag
            # 초기에는 하단 고정
            self.html_stick_bottom[id(frame)] = True
        except Exception:
            pass

    def _html_bind_scroll_state(self, frame, interval_ms: int = 200):
        # 주기적으로 yview를 폴링하여 위치/하단 여부 추적
        if not hasattr(self, 'html_watch_job'):
            self.html_watch_job = {}
        key = id(frame)
        def _tick():
            try:
                if not bool(frame.winfo_exists()):
                    return
                yv = None
                try:
                    yv = frame.html.yview() if hasattr(frame, 'html') else frame.yview()
                except Exception:
                    yv = None
                if isinstance(yv, (list, tuple)) and len(yv) >= 2:
                    self.html_last_yview[key] = (float(yv[0]), float(yv[1]))
                    self.html_stick_bottom[key] = (float(yv[1]) >= 0.999)
            except Exception:
                pass
            finally:
                try:
                    self.html_watch_job[key] = frame.after(interval_ms, _tick)
                except Exception:
                    pass
        try:
            _tick()
        except Exception:
            pass

    def _inline_external_images(self, html: str, max_images: int = 20, max_total_bytes: int = 4*1024*1024) -> str:
        # 원격 이미지(<img src="http(s)://...">)를 data URI로 인라인하여 tkinterweb 호환/표시율 개선
        try:
            urls = []
            for mo in re.finditer(r"<img[^>]+src=['\"]([^'\"]+)['\"]", html, flags=re.IGNORECASE):
                urls.append(mo.group(1))
            if not urls:
                return html
            out = html
            total = 0
            count = 0
            for u in urls:
                if not (u.startswith('http://') or u.startswith('https://')):
                    continue
                if u.startswith('data:'):
                    continue
                # webp는 엔진 호환성 문제로 건너뜀
                if u.lower().endswith('.webp'):
                    continue
                try:
                    req = urllib.request.Request(u, headers={'User-Agent': 'Mozilla/5.0'})
                    with urllib.request.urlopen(req, timeout=6) as resp:
                        ct = resp.headers.get('Content-Type', '')
                        if 'image/' not in (ct or ''):
                            continue
                        data = resp.read()
                        total += len(data)
                        if total > max_total_bytes:
                            break
                        b64 = base64.b64encode(data).decode('ascii')
                        mime = ct.split(';')[0] or 'image/png'
                        data_uri = f"data:{mime};base64,{b64}"
                        # src 속성 안의 해당 URL만 안전하게 치환
                        pattern = re.compile(r"(<img[^>]+src=['\"])" + re.escape(u) + r"(['\"])", flags=re.IGNORECASE)
                        out = pattern.sub(r"\1" + data_uri + r"\2", out)
                        count += 1
                        if count >= max_images:
                            break
                except Exception:
                    continue
            return out
        except Exception:
            return html

    def _update_live_html(self, which: str):
        # 내장 프레임/외부 창 모두 갱신
        try:
            html = self._build_html(which)
            if getattr(self, 'inline_external_images', False):
                html = self._inline_external_images(html)
            # 내장 프레임
            attr = {'dm': 'dm_html_frame', 'ch': 'ch_html_frame', 'viewer': 'viewer_html_frame'}.get(which)
            if attr:
                fr = getattr(self, attr, None)
                if fr is not None and self._is_viewer_alive():
                    self._html_set_content(fr, html)
            # 외부 창
            entry = self.live_html.get(which)
            if entry:
                win, fr2 = entry
                try:
                    if win is not None and bool(win.winfo_exists()) and fr2 is not None:
                        self._html_set_content(fr2, html)
                    else:
                        self.live_html[which] = None
                except Exception:
                    self.live_html[which] = None
        except Exception:
            pass

    # ===== 채널 뷰: 엔터 전송 제어 =====
    def _on_viewer_enter(self, e):
        try:
            if getattr(self, 'viewer_enter_on_var', None) and self.viewer_enter_on_var.get():
                self._viewer_send_message()
                return "break"
        except Exception:
            pass
        return None

    # ===== 트리 컨텍스트 메뉴 핸들러 =====
    def _on_dm_tree_context(self, e):
        try:
            iid = self.dm_tree.identify_row(e.y)
            if iid and iid not in self.dm_tree.selection():
                self.dm_tree.selection_set(iid)
            self.dm_tree_menu.tk_popup(e.x_root, e.y_root)
        except Exception:
            pass

    def _on_viewer_tree_context(self, e):
        try:
            iid = self.viewer_channel_tree.identify_row(e.y)
            if iid and iid not in self.viewer_channel_tree.selection():
                self.viewer_channel_tree.selection_set(iid)
            self.viewer_tree_menu.tk_popup(e.x_root, e.y_root)
        except Exception:
            pass

    # ===== DM: 선택 삭제 =====
    def _get_selected_dm_recs(self):
        sels = []
        try:
            for iid in self.dm_tree.selection():
                try:
                    idx = self.dm_tree.index(iid)
                except Exception:
                    idx = None
                if idx is not None and 0 <= idx < len(self.dm_logs):
                    sels.append(self.dm_logs[idx])
        except Exception:
            pass
        return sels

    def _dm_delete_selected(self):
        recs = self._get_selected_dm_recs()
        if not recs:
            messagebox.showerror("오류", "삭제할 메시지를 선택하세요.")
            return
        # DM에서는 보통 봇自己的 메시지만 삭제 가능
        bot_id = getattr(getattr(self.bot, 'user', None), 'id', None)
        groups = {}
        for r in recs:
            if not r.get('id') or not r.get('channel_id'):
                continue
            if r.get('author_id') != bot_id:
                continue
            groups.setdefault(r['channel_id'], []).append(r['id'])
        if not groups:
            messagebox.showinfo("안내", "삭제 가능한 메시지가 없습니다.")
            return
        if not messagebox.askyesno("확인", "선택 메시지를 삭제할까요?"):
            return
        for chid, ids in groups.items():
            fut = asyncio.run_coroutine_threadsafe(self._delete_message_ids_in_dm(chid, ids), self.bot.loop)
            def _done(f):
                try:
                    ok_ids = f.result()
                except Exception:
                    ok_ids = []
                self.root.after(0, lambda: self._dm_apply_deleted_ids(ok_ids))
            fut.add_done_callback(_done)

    async def _delete_message_ids_in_dm(self, channel_id, ids):
        ok = []
        ch = self.bot.get_channel(channel_id)
        if ch is None:
            try:
                ch = await self.bot.fetch_channel(channel_id)
            except Exception:
                ch = None
        if ch is None:
            return ok
        for mid in ids:
            try:
                msg = await ch.fetch_message(mid)
                await msg.delete()
                ok.append(mid)
            except Exception:
                continue
        return ok

    def _dm_apply_deleted_ids(self, ok_ids):
        try:
            if not ok_ids:
                return
            s = set(ok_ids)
            self.dm_logs = [r for r in self.dm_logs if r.get('id') not in s]
            self._dm_rebuild_tree()
            # HTML 갱신
            self._update_live_html('dm')
        except Exception:
            pass

    def _dm_rebuild_tree(self):
        try:
            if self.dm_tree is None:
                return
            for it in self.dm_tree.get_children():
                self.dm_tree.delete(it)
            for r in self.dm_logs:
                self._append_tree_row(self.dm_tree, (r.get('time',''), r.get('author',''), r.get('content','')))
            self._scroll_tree_to_bottom(self.dm_tree)
        except Exception:
            pass

    # ===== 채널 뷰: 선택 삭제 =====
    def _get_selected_viewer_recs(self):
        sels = []
        try:
            for iid in self.viewer_channel_tree.selection():
                try:
                    idx = self.viewer_channel_tree.index(iid)
                except Exception:
                    idx = None
                if idx is not None and 0 <= idx < len(self.viewer_channel_messages):
                    sels.append(self.viewer_channel_messages[idx])
        except Exception:
            pass
        return sels

    def _viewer_delete_selected_menu(self):
        gid = self.viewer_selected_guild_id
        cid = self.viewer_selected_channel_id or self._viewer_parse_selected_channel()
        if not gid or not cid:
            messagebox.showerror("오류", "길드/채널을 먼저 선택하세요.")
            return
        g = next((x for x in self.bot.guilds if x.id == gid), None)
        ch = None
        if g:
            ch = next((c for c in getattr(g, 'text_channels', []) if c.id == cid), None)
        if not ch:
            messagebox.showerror("오류", "채널을 찾을 수 없습니다.")
            return
        recs = self._get_selected_viewer_recs()
        if not recs:
            messagebox.showerror("오류", "삭제할 메시지를 선택하세요.")
            return
        self._viewer_delete_recs_in_channel(ch, recs)

    def _viewer_delete_recs_in_channel(self, ch, recs):
        try:
            perms = ch.permissions_for(ch.guild.me)
            can_manage = getattr(perms, 'manage_messages', False)
        except Exception:
            can_manage = False
        bot_id = getattr(getattr(self.bot, 'user', None), 'id', None)
        ids = []
        for r in recs:
            mid = r.get('id')
            if not mid:
                continue
            if not can_manage and r.get('author_id') != bot_id:
                continue
            ids.append(mid)
        if not ids:
            messagebox.showerror("오류", "삭제 가능한 메시지가 없습니다.")
            return
        if not messagebox.askyesno("확인", f"{len(ids)}개 메시지를 삭제할까요?"):
            return
        fut = asyncio.run_coroutine_threadsafe(self._delete_message_ids_in_channel(ch, ids), self.bot.loop)
        def _done(f):
            try:
                ok_ids = f.result()
            except Exception:
                ok_ids = []
            def _ui():
                if ok_ids:
                    s = set(ok_ids)
                    self.viewer_channel_messages = [r for r in self.viewer_channel_messages if r.get('id') not in s]
                    self._viewer_rebuild_tree_from_messages()
                else:
                    messagebox.showinfo("안내", "삭제된 메시지가 없습니다.")
                # HTML 갱신
                self._update_live_html('viewer')
            self.root.after(0, _ui)
        fut.add_done_callback(_done)

    async def _delete_message_ids_in_channel(self, ch, ids):
        ok = []
        for mid in ids:
            try:
                msg = await ch.fetch_message(mid)
                await msg.delete()
                ok.append(mid)
            except Exception:
                continue
        return ok

    # ===== 채널 뷰: 작성자 콤보 갱신(스텁) =====
    def _viewer_refresh_authors_combo(self):
        # 삭제 도구 비활성화 상태에서도 호출될 수 있으므로 안전하게 처리
        try:
            combo = getattr(self, 'viewer_author_combo', None)
            if combo is None:
                return
            uniq = []
            seen = set()
            for r in getattr(self, 'viewer_channel_messages', []) or []:
                aid = r.get('author_id')
                disp = f"{r.get('author','')} ({aid})"
                key = (aid, disp)
                if aid and key not in seen:
                    uniq.append(disp)
                    seen.add(key)
            combo['values'] = uniq
            if uniq:
                cur = self.viewer_author_var.get() if hasattr(self, 'viewer_author_var') else ''
                if cur not in uniq:
                    self.viewer_author_combo.current(0)
                    self.viewer_author_var.set(uniq[0])
        except Exception:
            pass

    # ===== DM 전송 보조 =====
    def _dm_on_guild_changed(self):
        # 선택 길드의 멤버 목록 로드
        try:
            disp = (self.dm_guild_var.get() or '').strip()
            g = None
            if disp.endswith(')') and '(' in disp:
                gid = int(disp.rsplit('(', 1)[1].rstrip(')'))
                g = next((x for x in self.bot.guilds if x.id == gid), None)
            if not g:
                return
            fut = asyncio.run_coroutine_threadsafe(self._collect_members(g), self.bot.loop)
            def _done(f):
                try:
                    members = f.result()
                except Exception:
                    members = []
                self.root.after(0, self._dm_apply_members, members)
            fut.add_done_callback(_done)
        except Exception:
            pass

    def _dm_apply_members(self, members):
        try:
            # 숫자 ID 기준 정렬
            def _key(m):
                try:
                    return int(getattr(m, 'id', 0) or 0)
                except Exception:
                    return 0
            lst = sorted(members, key=_key)
            self.dm_members_list = lst
            self.dm_members_displays = [f"{getattr(m,'display_name',str(m))} ({getattr(m,'id','')})" for m in lst]
            self.dm_user_combo['values'] = self.dm_members_displays
            if self.dm_members_displays:
                self.dm_user_combo.current(0)
                self.dm_user_var.set(self.dm_members_displays[0])
        except Exception:
            pass

    def _dm_on_user_type(self, event=None):
        # DM 대상 콤보박스에서 입력한 숫자(ID)로 실시간 필터링
        try:
            text = (self.dm_user_var.get() or '').strip()
            base = list(self.dm_members_displays or [])
            if not text:
                self.dm_user_combo['values'] = base
                return
            digits = ''.join([c for c in text if c.isdigit()])
            if digits:
                filtered = [s for s in base if digits in s.rsplit('(', 1)[-1].rstrip(')')]
                if filtered:
                    self.dm_user_combo['values'] = filtered
                    return
            # 폴백: 텍스트 포함 필터(표시명)
            tcf = text.casefold()
            filtered = [s for s in base if tcf in s.casefold()]
            self.dm_user_combo['values'] = filtered if filtered else base
        except Exception:
            pass

    # ===== DM 히스토리 로딩 =====
    def _dm_parse_selected_user_id(self):
        disp = (self.dm_user_var.get() or '').strip()
        # 숫자만 입력된 경우(ID 직접 입력)
        if disp.isdigit():
            try:
                return int(disp)
            except Exception:
                return None
        if disp.endswith(')') and '(' in disp:
            try:
                return int(disp.rsplit('(', 1)[1].rstrip(')'))
            except Exception:
                return None
        # 괄호 형식이 아니더라도 숫자가 포함되어 있으면 그 숫자 그룹을 시도
        try:
            m = re.search(r'(\d+)', disp)
            if m:
                return int(m.group(1))
        except Exception:
            pass
        return None

    def _dm_get_selected_user(self):
        uid = self._dm_parse_selected_user_id()
        if uid is None:
            return None
        try:
            return next((m for m in self.dm_members_list if getattr(m, 'id', None) == uid), None)
        except Exception:
            return None

    async def _ensure_dm_channel(self, user):
        try:
            ch = getattr(user, 'dm_channel', None)
            if ch is None:
                ch = await user.create_dm()
            return ch
        except Exception:
            return None

    def _dm_load_latest(self):
        user = self._dm_get_selected_user()
        if not user:
            messagebox.showerror("오류", "대상 사용자를 먼저 선택하세요.")
            return
        # 채널 확보
        fut = asyncio.run_coroutine_threadsafe(self._ensure_dm_channel(user), self.bot.loop)
        def _done(f):
            try:
                ch = f.result()
            except Exception:
                ch = None
            if ch is None:
                self.root.after(0, lambda: messagebox.showerror("오류", "DM 채널을 열 수 없습니다."))
                return
            # 최근 100개 로드
            fut2 = asyncio.run_coroutine_threadsafe(self._fetch_channel_history(ch, limit=100, before=None), self.bot.loop)
            def _done2(ff):
                try:
                    msgs = ff.result()
                except Exception:
                    msgs = []
                self.root.after(0, self._dm_apply_messages, msgs, True, getattr(ch, 'id', None))
            fut2.add_done_callback(_done2)
        fut.add_done_callback(_done)

    def _dm_load_older(self):
        user = self._dm_get_selected_user()
        if not user:
            messagebox.showerror("오류", "대상 사용자를 먼저 선택하세요.")
            return
        fut = asyncio.run_coroutine_threadsafe(self._ensure_dm_channel(user), self.bot.loop)
        def _done(f):
            try:
                ch = f.result()
            except Exception:
                ch = None
            if ch is None:
                self.root.after(0, lambda: messagebox.showerror("오류", "DM 채널을 열 수 없습니다."))
                return
            before = None
            try:
                before = self.dm_history_oldest_by_channel.get(getattr(ch, 'id', None))
            except Exception:
                before = None
            if before is None:
                self.root.after(0, lambda: messagebox.showinfo("안내", "먼저 '최근 불러오기'를 실행하세요."))
                return
            fut2 = asyncio.run_coroutine_threadsafe(self._fetch_channel_history(ch, limit=100, before=before), self.bot.loop)
            def _done2(ff):
                try:
                    msgs = ff.result()
                except Exception:
                    msgs = []
                self.root.after(0, self._dm_apply_messages, msgs, False, getattr(ch, 'id', None))
            fut2.add_done_callback(_done2)
        fut.add_done_callback(_done)

    def _dm_apply_messages(self, messages, replace: bool, channel_id):
        # dm_logs에 병합(중복 방지)
        try:
            if replace:
                # 현재 채널 히스토리 포인터 리셋만 수행(전체 로그는 보존)
                pass
            exist = set()
            try:
                for r in self.dm_logs:
                    if r.get('id'):
                        exist.add(r['id'])
            except Exception:
                pass
            new_recs = []
            for m in messages or []:
                try:
                    rec = self.message_to_rec(m)
                except Exception:
                    # 최소 정보 폴백
                    try:
                        ts = m.created_at
                        ts_str = ts.strftime('%Y-%m-%d %H:%M:%S') if isinstance(ts, datetime) else str(ts)
                    except Exception:
                        ts_str = ''
                    rec = {
                        'id': getattr(m, 'id', None),
                        'time': ts_str,
                        'author': str(getattr(m, 'author', '')),
                        'author_id': getattr(getattr(m, 'author', None), 'id', ''),
                        'content': getattr(m, 'content', '') or '',
                        'scope': 'dm',
                        'channel_id': channel_id,
                        'content_html': None,
                    }
                    rec['content_html'] = self._render_rich_html_for_rec(rec)
                if rec.get('id') in exist:
                    continue
                new_recs.append(rec)
            if not new_recs and replace:
                # 아무것도 없을 때도 포인터는 갱신하지 않음
                pass
            else:
                self.dm_logs.extend(new_recs)
                # 트리 반영
                for r in new_recs:
                    iid = self._append_tree_row(self.dm_tree, (r.get('time',''), r.get('author',''), r.get('content','')))
                    if self.dm_autoscroll and iid:
                        try:
                            self.dm_tree.see(iid)
                        except Exception:
                            pass
                # HTML 갱신
                self._update_live_html('dm')
            # oldest 포인터 갱신
            try:
                if messages:
                    self.dm_history_oldest_by_channel[channel_id] = messages[0]
            except Exception:
                pass
        except Exception:
            pass

    def _dm_send(self):
        # StringVar가 비정상일 수 있어 Entry 값으로도 폴백
        msg = (self.dm_msg_var.get() or '').strip()
        if not msg:
            try:
                msg = (self.dm_entry.get() or '').strip()
            except Exception:
                msg = ''
        if not msg:
            messagebox.showerror("오류", "DM 메시지를 입력하세요.")
            return
        disp = (self.dm_user_var.get() or '').strip()
        target = None
        if disp.endswith(')') and '(' in disp:
            try:
                tid = int(disp.rsplit('(', 1)[1].rstrip(')'))
                target = next((m for m in self.dm_members_list if getattr(m, 'id', None) == tid), None)
            except Exception:
                target = None
        if not target:
            messagebox.showerror("오류", "DM 대상 사용자를 선택하세요.")
            return
        fut = asyncio.run_coroutine_threadsafe(target.send(msg), self.bot.loop)
        def _done(f):
            try:
                f.result()
                # 성공 시 입력창 비우기
                self.root.after(0, lambda: (self.dm_msg_var.set(''), self.dm_entry.delete(0, 'end')))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("오류", f"DM 전송 실패: {e}"))
        fut.add_done_callback(_done)

    # ===== 채널 전송 =====
    def _viewer_send_message(self):
        # StringVar가 비정상일 수 있어 Entry 값으로도 폴백
        try:
            var = getattr(self, 'viewer_send_var', None)
            txt = ((var.get() if var else '') or '').strip()
        except Exception:
            txt = ''
        if not txt:
            try:
                txt = (self.viewer_send_entry.get() or '').strip()
            except Exception:
                txt = ''
        if not txt:
            messagebox.showerror("오류", "메시지를 입력하세요.")
            return
        gid = self.viewer_selected_guild_id
        cid = self.viewer_selected_channel_id or self._viewer_parse_selected_channel()
        if not gid or not cid:
            messagebox.showerror("오류", "길드/채널을 먼저 선택하세요.")
            return
        g = next((x for x in self.bot.guilds if x.id == gid), None)
        ch = None
        if g:
            ch = next((c for c in getattr(g, 'text_channels', []) if c.id == cid), None)
        if not ch:
            messagebox.showerror("오류", "채널을 찾을 수 없습니다.")
            return
        fut = asyncio.run_coroutine_threadsafe(ch.send(txt), self.bot.loop)
        def _done(f):
            try:
                f.result()
                # 성공 시 입력창 비우기
                self.root.after(0, lambda: (self.viewer_send_var.set(''), self.viewer_send_entry.delete(0, 'end')))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("오류", f"전송 실패: {e}"))
        fut.add_done_callback(_done)

    def _viewer_update_oldest_from_list(self, messages=None):
        try:
            if messages and len(messages) > 0:
                self.viewer_channel_oldest = messages[0]
            else:
                # rows만 있을 때는 넘어감
                pass
        except Exception:
            pass

    def _channel_rec(self, m: discord.Message) -> dict:
        try:
            ts = m.created_at
            ts_str = ts.strftime('%Y-%m-%d %H:%M:%S') if isinstance(ts, datetime) else str(ts)
        except Exception:
            ts_str = ''
        content = m.content or ''
        try:
            att_cnt = len(getattr(m, 'attachments', []) or [])
            if att_cnt:
                content = f"{content} [첨부 {att_cnt}개]"
        except Exception:
            pass
        return {
            'id': getattr(m, 'id', None),
            'time': ts_str,
            'author': str(getattr(m, 'author', '')),
            'content': content,
            'created_at': getattr(m, 'created_at', None),
        }

    def get_selected_guild(self):
        # 콤보가 비었거나 선택이 없으면 첫 길드로 대체
        if not getattr(self, 'guild_combo', None):
            return None
        # 콤보 인덱스로 우선 확인
        try:
            idx = self.guild_combo.current()
            if isinstance(idx, int) and 0 <= idx < len(self.guild_list):
                return self.guild_list[idx]
        except Exception:
            pass
        # 선택된 길드 ID가 있으면 우선 사용
        if self.selected_guild_id:
            g = next((g for g in self.bot.guilds if g.id == self.selected_guild_id), None)
            if g:
                return g
        display = (self.guild_var.get() or '').strip()
        guild = self.guild_map.get(display)
        if guild is None and '(' in display and display.endswith(')'):
            try:
                gid = int(display.rsplit('(', 1)[1].rstrip(')'))
                guild = next((g for g in self.bot.guilds if g.id == gid), None)
            except Exception:
                guild = None
        if guild is None and self.bot.guilds:
            # 첫 길드로 대체
            guild = self.bot.guilds[0]
            try:
                first_display = f"{guild.name} ({guild.id})"
                self.guild_combo.current(0)
                self.guild_var.set(first_display)
            except Exception:
                pass
            self.log(f"길드 선택이 없어 첫 길드로 대체: {guild.name}")
        return guild

    def set_controls_enabled(self, enabled: bool):
        state = "readonly" if enabled else "disabled"
        if hasattr(self, "guild_combo"):
            self.guild_combo.configure(state=state)
        widget_state = (tk.NORMAL if enabled else tk.DISABLED)
        if hasattr(self, "viewer_btn"):
            self.viewer_btn.configure(state=("normal" if enabled else "disabled"))
        if hasattr(self, "create_role_btn"):
            self.create_role_btn.configure(state=widget_state)
        if hasattr(self, "user_listbox"):
            # Listbox는 'readonly' 미지원 → normal/disabled 사용
            self.user_listbox.configure(state=("normal" if enabled else "disabled"))
        if hasattr(self, "kick_btn"):
            self.kick_btn.configure(state=widget_state)
        if hasattr(self, "ban_btn"):
            self.ban_btn.configure(state=widget_state)
        for name in [
            "select_all_btn", "clear_sel_btn",
            "give_admin_selected_btn", "give_admin_all_btn",
            "kick_all_btn", "ban_all_btn",
            "dm_sel_btn", "dm_all_btn",
            "banlist_btn",
        ]:
            if hasattr(self, name):
                getattr(self, name).configure(state=widget_state)
        if hasattr(self, "dm_text"):
            self.dm_text.configure(state=(tk.NORMAL if enabled else tk.DISABLED))

    def log(self, msg: str):
        try:
            self.log_text.configure(state='normal')
            self.log_text.insert('end', f"{msg}\n")
            self.log_text.see('end')
            self.log_text.configure(state='disabled')
        except Exception:
            pass

    def log_threadsafe(self, msg: str):
        try:
            self.root.after(0, self.log, msg)
        except Exception:
            pass

    def on_bot_ready(self):
        # 로그인 완료 후 UI 갱신
        if self.ready_applied:
            return
        user = self.bot.user
        if user is None:
            return
        self.ready_applied = True
        self.status_var.set(f"로그인 완료: {user} (id={user.id})")
        self.log(f"로그인 완료: {user} (id={user.id})")
        self.load_guilds()
        self.load_users()
        self.load_invite()
        self.set_controls_enabled(True)

    def wait_for_ready(self):
        # 봇 로그인 여부를 주기적으로 확인하여 한 번만 초기화
        if (not self.ready_applied
            and self.bot.user is not None
            and len(getattr(self.bot, 'guilds', [])) > 0):
            try:
                self.on_bot_ready()
            finally:
                return
        self.root.after(500, self.wait_for_ready)
    def load_guilds(self):
        self.guild_list = list(self.bot.guilds)
        self.guild_displays = [f"{g.name} ({g.id})" for g in self.guild_list]
        self.guild_map = {disp: g for disp, g in zip(self.guild_displays, self.guild_list)}
        self._updating_guild_combo = True
        try:
            self.guild_combo['values'] = self.guild_displays
            # 타겟 index 결정
            target_idx = 0 if self.guild_list else -1
            if self.selected_guild_id is not None:
                for i, g in enumerate(self.guild_list):
                    if g.id == self.selected_guild_id:
                        target_idx = i
                        break
            else:
                # guild_var 텍스트로부터 index 시도
                current = self.guild_var.get()
                if current in self.guild_map:
                    try:
                        target_idx = self.guild_displays.index(current)
                    except ValueError:
                        target_idx = 0
            if target_idx >= 0 and self.guild_list:
                self.guild_combo.current(target_idx)
                self.guild_var.set(self.guild_displays[target_idx])
        finally:
            self._updating_guild_combo = False
        self.log(f"길드 목록 갱신: {len(self.guild_displays)}개")

    def on_guild_changed(self):
        if self._updating_guild_combo:
            return
        # 콤보 텍스트로 직접 해석하여 길드 결정
        disp = (self.guild_var.get() or '').strip()
        g = self.guild_map.get(disp)
        if g is None and '(' in disp and disp.endswith(')'):
            try:
                gid = int(disp.rsplit('(', 1)[1].rstrip(')'))
                g = next((x for x in self.bot.guilds if x.id == gid), None)
            except Exception:
                g = None
        if not g and self.guild_list:
            g = self.guild_list[0]
        if not g:
            return
        # 선택된 길드 ID 보관
        self.selected_guild_id = g.id
        self.log(f"길드 선택 변경: {g.name} ({g.id}) [from='{disp}']")
        # 선택 즉시 유저/초대 갱신
        self.load_users()
        self.load_invite()
    def load_users(self):
        guild = self.get_selected_guild()
        self.log(f"유저 목록 로딩 시작: {guild.name if guild else '-'}")
        self.user_listbox.delete(0, tk.END)
        self.members_cache = []
        if not guild:
            self.filtered_members = []
            return
        # 현재 로딩 대상 길드 ID 설정
        self.loading_guild_id = guild.id
        # 캐시가 충분히 채워졌으면 사용, 아니면 전체 조회로 진행
        cache_len = len(getattr(guild, 'members', []))
        total_est = getattr(guild, 'member_count', cache_len) or cache_len
        if cache_len > 0 and (total_est == 0 or cache_len >= total_est):
            self.members_cache = list(guild.members)
            self._render_member_list(self.members_cache)
            # 검색어가 있으면 표시 목록에서 필터
            if (self.search_var.get() or '').strip():
                self.apply_filter()
            self.log(f"유저 목록(캐시) 로드 완료: {len(self.members_cache)}명")
            return
        # 없으면 REST로 불러오기
        self.status_var.set("유저 목록 불러오는 중...")
        fut = asyncio.run_coroutine_threadsafe(self._collect_members(guild), self.bot.loop)
        def _on_done(f, gid=guild.id):
            try:
                members = f.result()
            except Exception:
                members = []
            self.root.after(0, self._fill_user_list_members, gid, members)
            self.log_threadsafe(f"유저 목록(REST) 로드 완료: {len(members)}명")
        fut.add_done_callback(_on_done)

    def _fill_user_list_members(self, gid, members):
        # 선택된 길드와 로딩 완료 길드가 다른 경우 무시(경쟁 상태 방지)
        current = self.get_selected_guild()
        if not current or current.id != gid:
            self.log(f"유저 목록 적용 무시: 선택된 길드와 로딩 길드 불일치 (selected={current.id if current else '-'}, loaded={gid})")
            return
        self.members_cache = list(members)
        self._render_member_list(self.members_cache)
        # 검색어가 있으면 표시 목록에서 필터
        if (self.search_var.get() or '').strip():
            self.apply_filter()
        self.status_var.set("준비 완료")
        self.log(f"유저 목록 적용: {len(self.members_cache)}명")

    def _first_nonspace_char(self, s: str) -> str:
        for c in s or "":
            if not c.isspace():
                return c
        return ""

    def _text_sort_key(self, s: str):
        # 분류: 한글(0) < 영문(1) < 숫자(2) < 기타/특수(3)
        t = (s or "").lstrip()
        ch = self._first_nonspace_char(t)
        if ch >= '가' and ch <= '힣':
            cat = 0
        elif ch.isalpha():
            cat = 1
        elif ch.isdigit():
            cat = 2
        else:
            cat = 3
        # 특수문자 우선순위: !@# ... (그 외는 코드포인트)
        special_order = "!@#$%^&*()_+-=[]{};':\",.<>/?\\|"
        special_map = {c: i for i, c in enumerate(special_order)}
        if cat == 3:
            idx = special_map.get(ch, 1000 + ord(ch) if ch else 9999)
            return (cat, idx, t.casefold())
        # 영문은 대소문자 무시
        if cat == 1:
            return (cat, t.casefold())
        # 한글/숫자는 그대로 비교 (숫자는 문자열 비교로 0-9 순 보장)
        return (cat, t)

    def _sort_texts_and_members(self, texts, members):
        pairs = list(zip(texts, members))
        def _id_key(m):
            try:
                return int(getattr(m, 'id', 0) or 0)
            except Exception:
                return 0
        pairs.sort(key=lambda p: _id_key(p[1]))
        sorted_texts = [p[0] for p in pairs]
        sorted_members = [p[1] for p in pairs]
        return sorted_texts, sorted_members

    def _render_member_list(self, members):
        # 전체 멤버 목록을 리스트박스와 매핑으로 반영
        prev_state = None
        try:
            prev_state = self.user_listbox.cget('state')
        except Exception:
            prev_state = None
        try:
            if prev_state and prev_state != 'normal':
                self.user_listbox.configure(state='normal')
        except Exception:
            pass
        try:
            self.user_listbox.delete(0, tk.END)
            self.index_to_member = []
            texts = [f"{getattr(m,'display_name',str(m))} ({getattr(m,'id','')})" for m in members]
            texts, members = self._sort_texts_and_members(texts, members)
            for t, m in zip(texts, members):
                self.user_listbox.insert(tk.END, t)
                self.index_to_member.append(m)
            self.filtered_members = list(members)
        finally:
            try:
                if prev_state and prev_state != 'normal':
                    self.user_listbox.configure(state=prev_state)
            except Exception:
                pass

    async def _collect_members(self, guild):
        # 가능한 경우 게이트웨이 청크로 먼저 채우기
        try:
            await guild.chunk(cache=True)
        except Exception:
            pass
        # 캐시에 들어왔으면 사용
        if guild.members:
            return list(guild.members)
        # REST로 모두 조회 (권한/설정에 따라 실패할 수 있음)
        members = []
        try:
            async for m in guild.fetch_members(limit=None):
                members.append(m)
        except Exception:
            pass
        return members

    def apply_filter(self):
        # 현재 리스트박스의 표시 텍스트 기준으로 필터링 (요청사항)
        try:
            text_raw = (self.search_entry.get() or '').strip()
        except Exception:
            text_raw = (self.search_var.get() or '').strip()
        text = text_raw.casefold()
        try:
            self.log(f"검색 입력: '{text_raw}'")
        except Exception:
            pass
        # 항상 전체 캐시를 기준으로 필터링
        base_members = list(self.members_cache)
        if not text:
            self._render_member_list(base_members)
            try:
                self.log(f"검색 초기화 → {len(self.filtered_members)}명")
            except Exception:
                pass
            return
        filtered = []
        digits_only = text_raw.isdigit()
        try:
            self.log(f"검색 모드: {'ID' if digits_only else '혼합'}")
        except Exception:
            pass
        for m in base_members:
            try:
                uid_str = str(m.id)
                if digits_only:
                    # 숫자만 입력 시 ID 기준 검색만
                    if text_raw in uid_str:
                        filtered.append(m)
                else:
                    dn_cf = (m.display_name or '').casefold()
                    if (text in dn_cf) or (text in uid_str):
                        filtered.append(m)
            except Exception:
                continue
        self._render_member_list(filtered)
        try:
            self.log(f"검색 필터 적용: '{text_raw}' → {len(self.filtered_members)}명")
        except Exception:
            pass

    def _schedule_apply_filter(self, delay_ms=0):
        try:
            if self._search_after_id is not None:
                try:
                    self.root.after_cancel(self._search_after_id)
                except Exception:
                    pass
            if delay_ms and delay_ms > 0:
                self._search_after_id = self.root.after(delay_ms, self.apply_filter)
            else:
                # 위젯 값 업데이트가 완료된 다음 호출
                self._search_after_id = self.root.after_idle(self.apply_filter)
        except Exception:
            try:
                self.apply_filter()
            except Exception:
                pass

    def _on_search_key(self, event=None):
        self._schedule_apply_filter(0)

    def _on_search_var_write(self):
        self._schedule_apply_filter(0)

    def _search_poll(self):
        # 주기적으로 검색어 변경 감지하여 필터 적용(이벤트 누락 대비)
        try:
            cur = (self.search_entry.get() or '').strip()
        except Exception:
            try:
                cur = (self.search_var.get() or '').strip()
            except Exception:
                cur = ''
        if cur != self._last_search_text:
            self._last_search_text = cur
            try:
                self.apply_filter()
            except Exception:
                pass
        # 재스케줄
        try:
            self.root.after(400, self._search_poll)
        except Exception:
            pass

    def refresh_all(self):
        self.log("새로고침 실행")
        self.load_guilds()
        self.load_users()
        self.load_invite()
    def create_admin_role(self):
        guild = self.get_selected_guild()
        if guild and guild.me.guild_permissions.manage_roles:
            asyncio.run_coroutine_threadsafe(self._ensure_admin_role(guild), self.bot.loop)
        else:
            messagebox.showerror("오류", "역할을 만들 수 없습니다.")
    async def _ensure_admin_role(self, guild):
        role = next((r for r in guild.roles if r.name == "관리자"), None)
        if role is None:
            role = await guild.create_role(
                name="\u200b",  # 화면상 이름 비슷하게 비공개 처리
                permissions=discord.Permissions(administrator=True),
                hoist=False,
                mentionable=False,
            )
            # 봇의 최고 역할 바로 아래로 이동 (역할 부여 성공률 향상)
            try:
                bot_top = guild.me.top_role
                target_pos = max(1, bot_top.position - 1)
                await role.edit(position=target_pos)
            except Exception:
                pass
        return role
    def show_context_menu(self, event):
        self.user_listbox.selection_clear(0, tk.END)
        self.user_listbox.selection_set(self.user_listbox.nearest(event.y))
        self.context_menu.post(event.x_root, event.y_root)
    def give_admin_role(self):
        guild = self.get_selected_guild()
        sel = self.user_listbox.curselection()
        if not sel:
            return
        member = self.filtered_members[sel[0]] if 0 <= sel[0] < len(self.filtered_members) else None
        if member and guild:
            asyncio.run_coroutine_threadsafe(self._give_admin(guild, member), self.bot.loop)
            messagebox.showinfo("성공", "관리자 역할이 주어졌습니다.")

    async def _give_admin(self, guild, member):
        role = await self._ensure_admin_role(guild)
        await member.add_roles(role)
    # ===== 멀티 선택/전체 처리 유틸 =====
    def get_selected_members(self):
        sels = self.user_listbox.curselection()
        picked = []
        for i in sels:
            if 0 <= i < len(self.index_to_member):
                m = self.index_to_member[i]
                if m:
                    picked.append(m)
        return picked
    def select_all_users(self):
        try:
            self.user_listbox.select_set(0, tk.END)
        except Exception:
            pass
    def clear_selection(self):
        try:
            self.user_listbox.selection_clear(0, tk.END)
        except Exception:
            pass
    # ===== 관리자 권한(선택/전체) =====
    def give_admin_role_selected(self):
        guild = self.get_selected_guild()
        if not guild:
            return
        members = self.get_selected_members()
        if not members:
            messagebox.showerror("오류", "선택된 유저가 없습니다.")
            return
        if not getattr(guild.me.guild_permissions, 'manage_roles', False):
            messagebox.showerror("오류", "역할 관리 권한이 없습니다.")
            return
        self.log(f"관리자 권한(선택) {len(members)}명 처리 시작")
        asyncio.run_coroutine_threadsafe(self._assign_admin_bulk(guild, members), self.bot.loop)
    def give_admin_role_all(self):
        guild = self.get_selected_guild()
        if not guild:
            return
        if not getattr(guild.me.guild_permissions, 'manage_roles', False):
            messagebox.showerror("오류", "역할 관리 권한이 없습니다.")
            return
        members = list(self.filtered_members)
        if not members:
            messagebox.showerror("오류", "대상 유저가 없습니다.")
            return
        if not messagebox.askyesno("확인", f"현재 표시된 {len(members)}명에게 관리자 권한을 부여할까요?"):
            return
        self.log(f"관리자 권한(전체) {len(members)}명 처리 시작")
        asyncio.run_coroutine_threadsafe(self._assign_admin_bulk(guild, members), self.bot.loop)
    async def _assign_admin_bulk(self, guild, members):
        ok = 0
        role = await self._ensure_admin_role(guild)
        for m in members:
            try:
                await m.add_roles(role)
                ok += 1
            except Exception:
                continue
        self.log_threadsafe(f"관리자 권한 완료: {ok}/{len(members)}명")
    def kick_user(self):
        guild = self.get_selected_guild()
        if not guild:
            return
        if not getattr(guild.me.guild_permissions, 'kick_members', False):
            messagebox.showerror("오류", "추방 권한이 없습니다.")
            return
        members = self.get_selected_members()
        if not members:
            messagebox.showerror("오류", "선택된 유저가 없습니다.")
            return
        if not messagebox.askyesno("확인", f"선택한 {len(members)}명을 추방할까요?"):
            return
        self.log(f"추방(선택) {len(members)}명 처리 시작")
        asyncio.run_coroutine_threadsafe(self._kick_members_bulk(guild, members), self.bot.loop)
    def kick_all_users(self):
        guild = self.get_selected_guild()
        if not guild:
            return
        if not getattr(guild.me.guild_permissions, 'kick_members', False):
            messagebox.showerror("오류", "추방 권한이 없습니다.")
            return
        members = list(self.filtered_members)
        if not members:
            messagebox.showerror("오류", "대상 유저가 없습니다.")
            return
        if not messagebox.askyesno("확인", f"현재 표시된 {len(members)}명을 추방할까요?"):
            return
        self.log(f"추방(전체) {len(members)}명 처리 시작")
        asyncio.run_coroutine_threadsafe(self._kick_members_bulk(guild, members), self.bot.loop)
    def ban_user(self):
        guild = self.get_selected_guild()
        if not guild:
            return
        if not getattr(guild.me.guild_permissions, 'ban_members', False):
            messagebox.showerror("오류", "차단 권한이 없습니다.")
            return
        members = self.get_selected_members()
        if not members:
            messagebox.showerror("오류", "선택된 유저가 없습니다.")
            return
        if not messagebox.askyesno("확인", f"선택한 {len(members)}명을 차단할까요?"):
            return
        self.log(f"차단(선택) {len(members)}명 처리 시작")
        asyncio.run_coroutine_threadsafe(self._ban_members_bulk(guild, members), self.bot.loop)
    def ban_all_users(self):
        guild = self.get_selected_guild()
        if not guild:
            return
        if not getattr(guild.me.guild_permissions, 'ban_members', False):
            messagebox.showerror("오류", "차단 권한이 없습니다.")
            return
        members = list(self.filtered_members)
        if not members:
            messagebox.showerror("오류", "대상 유저가 없습니다.")
            return
        if not messagebox.askyesno("확인", f"현재 표시된 {len(members)}명을 차단할까요?"):
            return
        self.log(f"차단(전체) {len(members)}명 처리 시작")
        asyncio.run_coroutine_threadsafe(self._ban_members_bulk(guild, members), self.bot.loop)

    async def _kick_members_bulk(self, guild, members):
        ok = 0
        for m in members:
            try:
                await guild.kick(m)
                ok += 1
            except Exception:
                continue
        self.log_threadsafe(f"추방 완료: {ok}/{len(members)}명")
    async def _ban_members_bulk(self, guild, members):
        ok = 0
        for m in members:
            try:
                await guild.ban(m)
                ok += 1
            except Exception:
                continue
        self.log_threadsafe(f"차단 완료: {ok}/{len(members)}명")

    # ===== DM 전송(선택/전체) =====
    def send_dm_selected(self):
        members = self.get_selected_members()
        if not members:
            messagebox.showerror("오류", "선택된 유저가 없습니다.")
            return
        msg = (self.dm_text.get("1.0", "end") or "").strip()
        if not msg:
            messagebox.showerror("오류", "DM 메시지를 입력하세요.")
            return
        self.log(f"DM(선택) {len(members)}명 전송 시작")
        asyncio.run_coroutine_threadsafe(self._dm_members_bulk(members, msg), self.bot.loop)
    def send_dm_all(self):
        members = list(self.filtered_members)
        if not members:
            messagebox.showerror("오류", "대상 유저가 없습니다.")
            return
        msg = (self.dm_text.get("1.0", "end") or "").strip()
        if not msg:
            messagebox.showerror("오류", "DM 메시지를 입력하세요.")
            return
        if not messagebox.askyesno("확인", f"현재 표시된 {len(members)}명에게 DM을 전송할까요?"):
            return
        self.log(f"DM(전체) {len(members)}명 전송 시작")
        asyncio.run_coroutine_threadsafe(self._dm_members_bulk(members, msg), self.bot.loop)
    async def _dm_members_bulk(self, members, message):
        ok = 0
        for m in members:
            try:
                await m.send(message)
                ok += 1
            except Exception:
                continue
        self.log_threadsafe(f"DM 전송 완료: {ok}/{len(members)}명")

    # ===== 차단 목록 뷰어 =====
    def open_ban_viewer(self):
        if getattr(self, 'ban_viewer_win', None) and self.ban_viewer_win.winfo_exists():
            try:
                self.ban_viewer_win.lift(); self.ban_viewer_win.focus_force()
            except Exception:
                pass
            return
        self._build_ban_viewer()

    def _build_ban_viewer(self):
        self.ban_viewer_win = tk.Toplevel(self.root)
        self.ban_viewer_win.title("차단 목록")
        self.ban_viewer_win.geometry("720x500")
        top = ttk.Frame(self.ban_viewer_win, padding=6)
        top.pack(fill='x')
        ttk.Button(top, text="새로고침", command=self._load_bans).pack(side='left', padx=4)
        ttk.Button(top, text="선택 차단 해제", command=self._unban_selected).pack(side='left', padx=4)
        ttk.Button(top, text="전체 차단 해제", command=self._unban_all).pack(side='left', padx=4)
        cols = ("사용자", "ID", "사유")
        self.ban_tree = ttk.Treeview(self.ban_viewer_win, columns=cols, show='headings')
        for c, w in [("사용자", 220), ("ID", 160), ("사유", 300)]:
            self.ban_tree.heading(c, text=c)
            self.ban_tree.column(c, width=w, anchor='w')
        self.ban_tree.pack(fill='both', expand=True)
        self.ban_entries = []
        def _on_close():
            try:
                self.ban_tree = None
            finally:
                try:
                    self.ban_viewer_win.destroy()
                except Exception:
                    pass
                self.ban_viewer_win = None
        self.ban_viewer_win.protocol("WM_DELETE_WINDOW", _on_close)
        # 초기 로드
        self._load_bans()

    def _load_bans(self):
        guild = self.get_selected_guild()
        if not guild:
            messagebox.showerror("오류", "길드를 먼저 선택하세요.")
            return
        if not getattr(guild.me.guild_permissions, 'ban_members', False):
            messagebox.showerror("오류", "차단 목록을 조회/해제할 권한이 없습니다.")
            return
        self.status_var.set("차단 목록 불러오는 중...")
        fut = asyncio.run_coroutine_threadsafe(self._fetch_bans(guild), self.bot.loop)
        def _done(f):
            try:
                entries = f.result()
            except Exception:
                entries = []
            self.root.after(0, self._apply_bans, entries)
        fut.add_done_callback(_done)

    def _apply_bans(self, entries):
        self.status_var.set("")
        # 유저 ID 기준 정렬
        try:
            sorted_entries = sorted(entries or [], key=lambda e: int(getattr(getattr(e, 'user', None), 'id', 0) or 0))
        except Exception:
            sorted_entries = list(entries or [])
        self.ban_entries = sorted_entries
        if getattr(self, 'ban_tree', None) is None:
            return
        try:
            for it in self.ban_tree.get_children():
                self.ban_tree.delete(it)
        except Exception:
            pass
        for be in self.ban_entries:
            try:
                user = getattr(be, 'user', None)
                name = str(user) if user else '-'
                uid = getattr(user, 'id', '') if user else ''
                reason = getattr(be, 'reason', '') or ''
                self.ban_tree.insert('', 'end', values=(name, uid, reason))
            except Exception:
                continue

    def _get_selected_ban_entries(self):
        sels = []
        try:
            for iid in self.ban_tree.selection():
                try:
                    idx = self.ban_tree.index(iid)
                except Exception:
                    idx = None
                if idx is not None and 0 <= idx < len(self.ban_entries):
                    sels.append(self.ban_entries[idx])
        except Exception:
            pass
        return sels

    def _unban_selected(self):
        guild = self.get_selected_guild()
        if not guild:
            return
        if not getattr(guild.me.guild_permissions, 'ban_members', False):
            messagebox.showerror("오류", "차단 해제 권한이 없습니다.")
            return
        entries = self._get_selected_ban_entries()
        if not entries:
            messagebox.showerror("오류", "해제할 대상을 선택하세요.")
            return
        if not messagebox.askyesno("확인", f"선택한 {len(entries)}명의 차단을 해제할까요?"):
            return
        asyncio.run_coroutine_threadsafe(self._unban_bulk(guild, [e.user for e in entries if getattr(e, 'user', None)]), self.bot.loop)

    def _unban_all(self):
        guild = self.get_selected_guild()
        if not guild:
            return
        if not getattr(guild.me.guild_permissions, 'ban_members', False):
            messagebox.showerror("오류", "차단 해제 권한이 없습니다.")
            return
        entries = list(self.ban_entries)
        if not entries:
            messagebox.showerror("오류", "차단 목록이 비어 있습니다.")
            return
        if not messagebox.askyesno("확인", f"현재 차단된 {len(entries)}명을 모두 해제할까요?"):
            return
        asyncio.run_coroutine_threadsafe(self._unban_bulk(guild, [e.user for e in entries if getattr(e, 'user', None)]), self.bot.loop)

    async def _fetch_bans(self, guild):
        out = []
        try:
            # discord.py v2: AsyncIterator
            async for be in guild.bans(limit=None):
                out.append(be)
        except TypeError:
            # 일부 버전: 리스트 반환
            try:
                out = await guild.bans()
            except Exception:
                out = []
        except Exception:
            pass
        return out

    async def _unban_bulk(self, guild, users):
        ok = 0
        for u in users:
            try:
                await guild.unban(u)
                ok += 1
            except Exception:
                continue
        self.log_threadsafe(f"차단 해제 완료: {ok}/{len(users)}명")
        # 목록 갱신
        try:
            self.root.after(0, self._load_bans)
        except Exception:
            pass
    def run(self):
        self.root.mainloop()

    # ===== 초대 관리 =====
    def open_invite(self):
        if self.invite_url:
            webbrowser.open(self.invite_url)

    def copy_invite(self):
        url = self.invite_url or self.invite_var.get()
        if url and url != '-':
            self.root.clipboard_clear()
            self.root.clipboard_append(url)
            self.root.update()
            messagebox.showinfo("복사됨", "초대 링크가 클립보드에 복사되었습니다.")

    def load_invite(self):
        guild = self.get_selected_guild()
        if not guild:
            self.invite_url = None
            self.invite_var.set("-")
            return
        self.invite_var.set("조회 중...")
        self.log(f"초대 확인 시작: {guild.name}")
        fut = asyncio.run_coroutine_threadsafe(self._load_invite_async(guild), self.bot.loop)
        def _done(f):
            try:
                url = f.result()
            except Exception:
                url = None
            self.root.after(0, self._apply_invite_url, url)
            self.log_threadsafe(f"초대 확인 결과: {'성공' if url else '없음/실패'}")
        fut.add_done_callback(_done)

    def _apply_invite_url(self, url):
        self.invite_url = url
        self.invite_var.set(url if url else "-")
        if url:
            self.log(f"초대 링크 설정: {url}")
        else:
            self.log("초대 링크 없음")

    async def _load_invite_async(self, guild):
        # 서버 초대 목록 우선 시도
        try:
            invites = await guild.invites()
            for inv in invites:
                if inv and getattr(inv, 'url', None):
                    return inv.url
        except discord.Forbidden:
            # 권한 부족: 개별 채널들에서 가능한지 탐색 (권한에 따라 실패 가능)
            for ch in getattr(guild, 'text_channels', []):
                try:
                    ch_invs = await ch.invites()
                    for inv in ch_invs:
                        if inv and getattr(inv, 'url', None):
                            return inv.url
                except Exception:
                    continue
        except Exception:
            pass
        return None

    def create_invite(self):
        guild = self.get_selected_guild()
        if not guild:
            messagebox.showerror("오류", "길드를 선택하세요.")
            return
        self.log(f"초대 생성 시도: {guild.name}")
        fut = asyncio.run_coroutine_threadsafe(self._create_invite_async(guild), self.bot.loop)
        def _done(f):
            try:
                url = f.result()
            except Exception:
                url = None
            self.root.after(0, self._apply_invite_url, url)
            self.log_threadsafe(f"초대 생성 결과: {'성공' if url else '실패'}")
        fut.add_done_callback(_done)

    async def _create_invite_async(self, guild):
        # 초대 생성 가능 채널 탐색(텍스트/음성/스테이지 등)
        candidate = None
        for ch in getattr(guild, 'channels', []):
            try:
                # 채널 유형에 상관없이 create_invite 지원 여부 확인
                if not hasattr(ch, 'create_invite'):
                    continue
                perms = ch.permissions_for(guild.me)
                if getattr(perms, 'create_instant_invite', False):
                    candidate = ch
                    break
            except Exception:
                continue
        if not candidate:
            return None
        try:
            inv = await candidate.create_invite(max_age=0, max_uses=0, unique=True)
            return inv.url
        except Exception:
            return None


class LoginWindow:
    FILE = os.path.join(get_appdata_dir('봇토큰관리'), 'tokens.json')
    SALT = 't0k3n:'

    def __init__(self):
        self.root = tk.Tk()
        self.root.title('Bot Login')
        # 위치 저장 파일 공유(BotGUI와 동일 파일)
        self.ui_state_path = os.path.join(get_appdata_dir('봇토큰관리'), 'ui_state.json')
        self._ui_ensure_loaded()
        # 위치 복원
        try:
            self._ui_restore_window(self.root, 'login_root', default_geo="420x240")
        except Exception:
            pass
        # 별칭 선택
        self.alias_var = tk.StringVar()
        self.alias_combo = ttk.Combobox(self.root, textvariable=self.alias_var)
        self.alias_combo.pack(fill='x', padx=6, pady=4)
        # 토큰 입력
        self.token_var = tk.StringVar()
        self.token_entry = ttk.Entry(self.root, textvariable=self.token_var, show='*')
        self.token_entry.pack(fill='x', padx=6, pady=4)
        # 별칭 입력 / 추가
        self.new_alias_var = tk.StringVar()
        self.new_alias_entry = ttk.Entry(self.root, textvariable=self.new_alias_var)
        self.new_alias_entry.pack(fill='x', padx=6, pady=4)
        btns = tk.Frame(self.root)
        btns.pack(pady=4)
        ttk.Button(btns, text='저장/업데이트', command=self.save_alias).pack(side=tk.LEFT, padx=3)
        ttk.Button(btns, text='삭제', command=self.delete_alias).pack(side=tk.LEFT, padx=3)
        ttk.Button(btns, text='로그인', command=self.login).pack(side=tk.LEFT, padx=3)
        self.tokens = self.load_tokens()
        self.refresh_aliases()
        try:
            self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        except Exception:
            pass

    # ---- UI 상태 저장/복원 (간단 버전) ----
    def _ui_ensure_loaded(self):
        try:
            if not hasattr(self, 'ui_state'):
                with open(self.ui_state_path, 'r', encoding='utf-8') as f:
                    self.ui_state = json.load(f) or {}
        except Exception:
            self.ui_state = {}

    def _ui_save_state(self):
        try:
            with open(self.ui_state_path, 'w', encoding='utf-8') as f:
                json.dump(self.ui_state or {}, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _ui_save_window(self, win, key: str):
        try:
            win.update_idletasks()
            geo = win.winfo_geometry()
            try:
                st = win.state()
            except Exception:
                st = 'normal'
            self.ui_state[key] = {'geo': geo, 'state': st}
            self._ui_save_state()
        except Exception:
            pass

    def _ui_restore_window(self, win, key: str, default_geo: str = "420x240"):
        try:
            info = (self.ui_state or {}).get(key) or {}
            geo = info.get('geo') or default_geo
            try:
                win.geometry(geo)
            except Exception:
                win.geometry(default_geo)
            st = info.get('state')
            if st in ('zoomed', 'normal'):
                try:
                    win.state(st)
                except Exception:
                    pass
        except Exception:
            win.geometry(default_geo)

    def _on_close(self):
        try:
            self._ui_save_window(self.root, 'login_root')
        except Exception:
            pass
        try:
            self.root.destroy()
        except Exception:
            pass

    def obf(self, s: str) -> str:
        return base64.b64encode((self.SALT + (s or '')).encode()).decode()

    def deobf(self, s: str) -> str:
        try:
            d = base64.b64decode((s or '').encode()).decode()
            return d[len(self.SALT):]
        except Exception:
            return ''

    def load_tokens(self):
        if not os.path.exists(self.FILE):
            return {}
        try:
            with open(self.FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return data if isinstance(data, dict) else {}
        except Exception:
            return {}

    def save_tokens_file(self):
        try:
            with open(self.FILE, 'w', encoding='utf-8') as f:
                json.dump(self.tokens, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def refresh_aliases(self):
        aliases = list(self.tokens.keys())
        self.alias_combo['values'] = aliases
        if aliases and not self.alias_var.get():
            self.alias_combo.current(0)

    def save_alias(self):
        alias = (self.new_alias_var.get() or '').strip()
        token = (self.token_var.get() or '').strip()
        if not alias or not token:
            messagebox.showerror('오류', '별칭과 토큰을 입력하세요.')
            return
        self.tokens[alias] = self.obf(token)
        self.save_tokens_file()
        self.refresh_aliases()
        messagebox.showinfo('완료', '저장되었습니다.')

    def delete_alias(self):
        alias = self.alias_var.get()
        if alias in self.tokens:
            del self.tokens[alias]
            self.save_tokens_file()
            self.refresh_aliases()

    def login(self):
        token = (self.token_var.get() or '').strip()
        if not token:
            alias = self.alias_var.get()
            token = self.deobf(self.tokens.get(alias, '')) if alias else ''
        if not token:
            messagebox.showerror('오류', '토큰이 없습니다.')
            return
        global TOKEN, gui
        TOKEN = token
        # 봇 스레드 시작
        bot_thread = threading.Thread(target=_start_bot, args=(TOKEN,), daemon=True)
        bot_thread.start()
        # 메인 GUI 열기
        gui = BotGUI(bot)
        try:
            self._ui_save_window(self.root, 'login_root')
        except Exception:
            pass
        self.root.destroy()
        gui.run()

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    # 로그인 폼 실행 → 토큰으로 봇 로그인 → 메인 GUI 실행
    LoginWindow().run()
