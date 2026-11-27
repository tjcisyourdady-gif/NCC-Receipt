"""自动更新提示弹窗、手动更新窗口及主界面更新控制器（后台线程查包，主线程安全刷新）。"""
import queue
import threading
import tkinter as tk
from tkinter import ttk

import receipt_release_identity as identity


class StartupUpdatePrompt(tk.Toplevel):
    """每次启动发现新版本后在前台提示一次；关闭或忽略仅对本次启动生效。"""

    def __init__(self, controller, data):
        super().__init__(controller.root)
        self.controller = controller
        self.title('发现新版本')
        self.geometry('560x350')
        self.minsize(480, 300)
        self.transient(controller.root)
        self.protocol('WM_DELETE_WINDOW', self.ignore)

        area = ttk.Frame(self, padding=18)
        area.pack(fill='both', expand=True)
        ttk.Label(
            area,
            text='发现收款结算单工具新版本 V' + data['latest_version'],
            font=('微软雅黑', 13, 'bold'),
        ).pack(anchor='w')
        ttk.Label(
            area,
            text='当前版本：V' + data['current_version'],
            font=('微软雅黑', 9),
        ).pack(anchor='w', pady=(6, 10))

        notes = tk.Text(area, height=7, wrap='word', font=('微软雅黑', 9))
        notes.insert('1.0', data.get('notes') or '有新的程序版本可用。')
        notes.configure(state='disabled')
        notes.pack(fill='both', expand=True)

        ttk.Label(
            area,
            text=(
                '更新后将自动重启，原有底稿模板、对照表方案、填写顺序设置与时间设置均会保留。'
                '点击“忽略”仅对本次启动生效。'
            ),
            wraplength=510,
            foreground='#555555',
            font=('微软雅黑', 9),
        ).pack(fill='x', pady=10)

        buttons = ttk.Frame(area)
        buttons.pack(fill='x')
        self.ignore_button = ttk.Button(buttons, text='忽略', command=self.ignore)
        self.ignore_button.pack(side='right')
        self.update_button = ttk.Button(buttons, text='立即更新', command=self.update_now)
        self.update_button.pack(side='right', padx=8)

        self.update_idletasks()
        if self.winfo_viewable():
            self.grab_set()
            self.update_button.focus_set()

    def ignore(self):
        self.controller.startup_prompt_handled = True
        self.controller.startup_pending_update = None
        self.destroy()

    def update_now(self):
        self.destroy()
        self.controller.start_prompted_update()

    def destroy(self):
        if getattr(self.controller, 'startup_update_dialog', None) is self:
            self.controller.startup_update_dialog = None
        super().destroy()


class UpdateDialog(tk.Toplevel):
    """手动检查、下载进度展示及安全安装重启窗口。"""

    def __init__(self, controller):
        super().__init__(controller.root)
        self.controller = controller
        self.title('收款结算单工具 - 自动更新')
        self.geometry('640x420')
        self.minsize(540, 350)
        self.transient(controller.root)

        outer = ttk.Frame(self, padding=16)
        outer.pack(fill='both', expand=True)

        self.version_var = tk.StringVar()
        self.message_var = tk.StringVar()
        ttk.Label(outer, textvariable=self.version_var, font=('微软雅黑', 10, 'bold')).pack(anchor='w')
        ttk.Label(
            outer,
            text='每次打开程序自动在后台检测更新包；优先国内直连通道，备用 GitHub 官方通道。',
            foreground='#555555',
            font=('微软雅黑', 9),
        ).pack(anchor='w', pady=(4, 6))
        ttk.Label(
            outer,
            textvariable=self.message_var,
            wraplength=590,
            font=('微软雅黑', 9),
        ).pack(fill='x', pady=(2, 8))

        self.progress = ttk.Progressbar(outer, maximum=100)
        self.progress.pack(fill='x')

        self.notes = tk.Text(outer, height=9, wrap='word', state='disabled', font=('微软雅黑', 9))
        self.notes.pack(fill='both', expand=True, pady=12)

        buttons = ttk.Frame(outer)
        buttons.pack(fill='x')
        self.check_button = ttk.Button(buttons, text='重新检查', command=controller.check_update)
        self.check_button.pack(side='left')
        self.download_button = ttk.Button(buttons, text='下载更新', command=self.download)
        self.download_button.pack(side='left', padx=8)
        self.install_button = ttk.Button(buttons, text='安装并重启', command=controller.install_update)
        self.install_button.pack(side='left')
        ttk.Button(buttons, text='关闭', command=self.destroy).pack(side='right')

        self.flow_hint_var = tk.StringVar()
        ttk.Label(
            outer,
            textvariable=self.flow_hint_var,
            wraplength=600,
            foreground='#555555',
            font=('微软雅黑', 9),
        ).pack(fill='x', pady=(10, 0))

        self.last_notes = None
        self.timer = None
        self.refresh()

    def download(self):
        threading.Thread(
            target=self.controller.updater.download,
            name='NCCReceiptUpdateDownload',
            daemon=True,
        ).start()

    def refresh(self):
        data = self.controller.updater.status()
        self.version_var.set(
            f"当前版本：V{data['current_version']}    检测到的版本："
            + (f"V{data['latest_version']}" if data['latest_version'] else '—')
        )
        source_suffix = f"（{data['source']}）" if data.get('source') else ''
        self.message_var.set(data['message'] + source_suffix)
        self.progress['value'] = data['progress']
        if data['notes'] != self.last_notes:
            self.notes.configure(state='normal')
            self.notes.delete('1.0', 'end')
            self.notes.insert('end', data['notes'] or '暂无新版本说明。')
            self.notes.configure(state='disabled')
            self.last_notes = data['notes']
        busy = self.controller.update_installing or data['status'] in (
            'checking', 'downloading', 'staging', 'starting', 'installing'
        )
        self.flow_hint_var.set(
            '下载完成后将自动安装并重启，请稍候。已有底稿模板、对照表和配置会完整保留。'
            if self.controller.auto_update_requested
            else '安装前请先停止填报任务。已有底稿模板、对照表、填写顺序设置及时间设置均会保留。'
        )
        self.check_button.configure(state='disabled' if busy else 'normal')
        self.download_button.configure(
            state='normal' if data['status'] == 'available' and not busy else 'disabled'
        )
        self.install_button.configure(
            state='normal'
            if data['status'] == 'ready' and not busy and not self.controller.is_busy()
            else 'disabled'
        )
        self.timer = self.after(300, self.refresh)

    def destroy(self):
        if self.timer is not None:
            self.after_cancel(self.timer)
            self.timer = None
        if getattr(self.controller, 'update_dialog', None) is self:
            self.controller.update_dialog = None
        super().destroy()


class ReceiptUpdateController:
    """衔接主窗口与 UpdateManager：负责启动自动检测、延后弹窗提示与安全安装重启。"""

    def __init__(self, root, updater, is_busy, on_lock_change, on_request_exit, status_var):
        self.root = root
        self.updater = updater
        self.is_busy = is_busy
        self.on_lock_change = on_lock_change
        self.on_request_exit = on_request_exit
        self.status_var = status_var
        self.events = queue.Queue()
        self.update_dialog = None
        self.startup_update_dialog = None
        self.startup_check_started = False
        self.startup_prompt_handled = False
        self.startup_pending_update = None
        self.auto_update_requested = False
        self.update_installing = False
        self.closing = False
        self._startup_timer_id = None
        self._poll_timer_id = None

    def schedule(self):
        self._startup_timer_id = self.root.after(1200, self.startup_check)
        self._poll_timer_id = self.root.after(400, self.poll)

    def cancel_timers(self):
        self.closing = True
        for attr in ('_startup_timer_id', '_poll_timer_id'):
            timer_id = getattr(self, attr, None)
            if timer_id is not None:
                try:
                    self.root.after_cancel(timer_id)
                except tk.TclError:
                    pass
                setattr(self, attr, None)

    def startup_check(self):
        self._startup_timer_id = None
        if self.closing or self.startup_check_started:
            return
        self.startup_check_started = True
        state = self.updater.status()
        if state['status'] == 'idle':
            self.check_update(startup=True)
        elif state['status'] == 'available':
            self._handle_startup_result(state)

    def check_update(self, startup=False):
        if self.update_installing:
            return
        if not startup:
            self.startup_prompt_handled = True
            self.startup_pending_update = None
        threading.Thread(
            target=self._check_worker,
            args=(startup,),
            name='NCCReceiptUpdateCheck',
            daemon=True,
        ).start()

    def _check_worker(self, startup):
        data = self.updater.check()
        if startup:
            self.events.put(('startup_checked', data))

    def _handle_startup_result(self, data):
        if self.startup_prompt_handled:
            return
        if data['status'] == 'available':
            self.startup_pending_update = data
        else:
            self.startup_prompt_handled = True

    def _can_show_startup_prompt(self):
        if self.closing or self.update_installing or self.is_busy():
            return False
        try:
            focus = self.root.focus_displayof()
            return (
                self.root.state() not in ('iconic', 'withdrawn')
                and self.root.winfo_viewable()
                and focus is not None
                and focus.winfo_toplevel() is self.root
                and self.root.grab_current() is None
            )
        except tk.TclError:
            return False

    def _maybe_show_startup_prompt(self):
        data = self.startup_pending_update
        if data is None or self.startup_prompt_handled:
            return
        current = self.updater.status()
        if current['status'] != 'available' or current['latest_version'] != data['latest_version']:
            self.startup_pending_update = None
            self.startup_prompt_handled = True
            return
        if not self._can_show_startup_prompt():
            return
        self.startup_prompt_handled = True
        self.startup_pending_update = None
        self.startup_update_dialog = StartupUpdatePrompt(self, data)

    def open_update_dialog(self):
        self.startup_prompt_handled = True
        self.startup_pending_update = None
        if self.update_dialog is not None and self.update_dialog.winfo_exists():
            self.update_dialog.lift()
            self.update_dialog.focus_force()
        else:
            self.update_dialog = UpdateDialog(self)

    def start_prompted_update(self):
        if self.update_installing:
            return
        if self.is_busy():
            self.updater._state(message='任务正在运行，请停止后安装更新。')
            self.open_update_dialog()
            return
        if self.updater.status()['status'] != 'available':
            self.open_update_dialog()
            return
        self.startup_prompt_handled = True
        self.startup_pending_update = None
        self.auto_update_requested = True
        self.update_installing = True
        self.on_lock_change(True)
        self.open_update_dialog()
        threading.Thread(
            target=self._download_prompted_worker,
            name='NCCReceiptStartupUpdate',
            daemon=True,
        ).start()

    def _download_prompted_worker(self):
        try:
            data = self.updater.download()
        except Exception as exc:
            data = self.updater._state(status='error', message=str(exc))
        kind = 'prompted_downloaded' if data['status'] == 'ready' else 'prompted_failed'
        self.events.put((kind, data))

    def install_update(self):
        if self.update_installing:
            return
        if self.is_busy():
            self.updater._state(message='请先停止当前任务，再安装更新。')
            return
        self.update_installing = True
        self.on_lock_change(True)
        self.updater._state(status='starting', message='正在准备安装，请稍候…')
        threading.Thread(
            target=self._install_worker,
            name='NCCReceiptUpdateInstall',
            daemon=False,
        ).start()

    def _install_worker(self):
        try:
            self.updater.start_installer(self.is_busy)
        except Exception as exc:
            self.events.put(('install_failed', str(exc)))
        else:
            self.events.put(('install_ready', self.updater.status()))

    def poll(self):
        self._poll_timer_id = None
        if self.closing:
            return
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == 'startup_checked':
                    self._handle_startup_result(payload)
                elif kind == 'prompted_downloaded':
                    self.update_installing = False
                    if self.is_busy():
                        self.auto_update_requested = False
                        self.on_lock_change(False)
                        self.updater._state(message='更新已下载，请停止任务后点击“安装并重启”。')
                    else:
                        self.install_update()
                elif kind == 'prompted_failed':
                    self.update_installing = False
                    self.auto_update_requested = False
                    self.on_lock_change(False)
                    print('自动更新下载失败：' + payload.get('message', ''))
                elif kind == 'install_ready':
                    self.cancel_timers()
                    self.on_request_exit()
                    return
                elif kind == 'install_failed':
                    self.update_installing = False
                    self.auto_update_requested = False
                    self.updater._state(
                        status='ready' if self.updater.ready is not None else 'error',
                        message=payload,
                    )
                    self.on_lock_change(False)
                    print('更新安装助手启动失败：' + str(payload))
        except queue.Empty:
            pass

        state = self.updater.status()
        self.status_var.set(f'V{identity.APP_VERSION} | {state["message"]}')
        self._maybe_show_startup_prompt()
        self._poll_timer_id = self.root.after(400, self.poll)
