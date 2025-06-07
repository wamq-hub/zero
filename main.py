import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import pandas as pd
import os
from datetime import datetime
import json


def create_sample_excel_files():
    """Create sample CSV files if none exist."""
    data = {
        'instructors.csv': [['1001', 'Alice', 'Math'], ['1002', 'Bob', 'CS']],
        'permissions.csv': [['1001', 'admin'], ['1002', 'instructor']],
        'students.csv': [['2001', 'Student1'], ['2002', 'Student2']],
        'schedule1.csv': [['2001', 'Math'], ['2002', 'CS']],
        'instructor_schedule.csv': [['1001', 'Math'], ['1002', 'CS']],
    }
    for fname, rows in data.items():
        if not os.path.exists(fname):
            with open(fname, 'w', encoding='utf-8') as f:
                for row in rows:
                    f.write(','.join(row) + '\n')


class TrainingManagementSystem:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("نظام إدارة التدريب والحرمان - مبادرة كامل (المحسن)")
        self.root.geometry('1200x800')
        self.root.configure(bg='#f0f0f0')

        self.instructors_data = {}
        self.permissions_data = {}
        self.student_schedule_data = pd.DataFrame()
        self.students_info_data = pd.DataFrame()
        self.instructor_schedule_data = pd.DataFrame()
        self.reports_data = []
        self.data_load_status = {}
        self.current_user = None
        self.selected_course = None
        self.selected_student = None
        self.data_integrity_issues = []

        self.load_excel_files()
        self.validate_data_integrity()
        self.create_login_interface()

    def read_file(self, filename):
        try:
            if filename.lower().endswith('.csv'):
                try:
                    df = pd.read_csv(filename, encoding='utf-8', sep=None, engine='python')
                except Exception:
                    df = pd.read_csv(filename, encoding='windows-1256', sep=None, engine='python')
            else:
                df = pd.read_excel(filename)
            df.columns = [str(c).strip() for c in df.columns]
            return df
        except Exception:
            return pd.DataFrame()

    def load_excel_files(self):
        files = [f for f in os.listdir('.') if f.endswith(('.csv', '.xlsx'))]
        instr_file = self.find_file_by_pattern(['instructor', 'trainer'], files)
        perm_file = self.find_file_by_pattern(['permission'], files)
        schedule_files = [f for f in files if 'schedule' in f and 'instructor' not in f]
        students_file = self.find_file_by_pattern(['student'], files)
        instr_sched_file = self.find_file_by_pattern(['instructor_schedule'], files)

        if instr_file:
            self.load_instructors_data(instr_file)
        else:
            self.create_sample_instructors()
            self.data_load_status['sample_instructors'] = True
        if perm_file:
            self.load_permissions_data(perm_file)
        else:
            self.create_sample_permissions()
            self.data_load_status['sample_permissions'] = True
        if schedule_files:
            self.load_schedule_data(schedule_files)
        else:
            self.create_sample_schedules()
            self.data_load_status['sample_schedules'] = True
        if students_file:
            self.load_students_data(students_file)
        else:
            self.create_sample_students()
            self.data_load_status['sample_students'] = True
        if instr_sched_file:
            self.load_instructor_schedule_data(instr_sched_file)
        else:
            self.create_sample_instructor_schedule()
            self.data_load_status['sample_instructor_schedule'] = True

        self.load_saved_reports()

    def validate_data_integrity(self):
        issues = []
        auto_ids = [i for i in self.instructors_data if str(i).startswith('AUTO_')]
        if len(auto_ids) != len(set(auto_ids)):
            issues.append('Duplicate AUTO_ IDs')
        if not self.student_schedule_data.empty:
            instr_names = {v['name'] for v in self.instructors_data.values()}
            sch_instructors = set(self.student_schedule_data.get('Instructor', []))
            diff = sch_instructors - instr_names
            if diff:
                issues.append('Unknown instructors in schedules: ' + ','.join(diff))
        if not self.student_schedule_data.empty and not self.students_info_data.empty:
            stu_ids = set(self.students_info_data.get('ID', []))
            sch_ids = set(self.student_schedule_data.get('StudentID', []))
            diff = sch_ids - stu_ids
            if diff:
                issues.append('Unknown students in schedules: ' + ','.join(map(str, diff)))
        self.data_integrity_issues = issues

    def find_file_by_pattern(self, patterns, files):
        for p in patterns:
            for f in files:
                if p.lower() in f.lower():
                    return f
        return None

    def load_instructors_data(self, filename):
        df = self.read_file(filename)
        if df.empty:
            self.create_sample_instructors()
            self.data_load_status['sample_instructors'] = True
            return
        for idx, row in df.iterrows():
            emp_id = str(row.iloc[0]) if pd.notna(row.iloc[0]) else f"AUTO_{idx:04d}"
            name = str(row.iloc[1]) if len(row) > 1 else f"Name{idx}"
            dept = str(row.iloc[2]) if len(row) > 2 else 'عام'
            self.instructors_data[emp_id] = {'name': name, 'dept': dept, 'role': 'instructor'}

    def create_sample_instructors(self):
        self.instructors_data = {
            'AUTO_0001': {'name': 'مدرب1', 'dept': 'عام', 'role': 'instructor'},
            'AUTO_0002': {'name': 'مدرب2', 'dept': 'عام', 'role': 'instructor'},
        }

    def load_permissions_data(self, filename):
        df = self.read_file(filename)
        if df.empty:
            self.create_sample_permissions()
            self.data_load_status['sample_permissions'] = True
            return
        for idx, row in df.iterrows():
            name = str(row.get('name', row.iloc[0]))
            role = str(row.get('role', row.iloc[-1])).lower()
            emp_id = str(row.get('emp_id', row.iloc[1] if len(row) > 1 else f"AUTO_{idx:04d}"))
            if emp_id not in self.instructors_data:
                self.instructors_data[emp_id] = {'name': name, 'dept': 'عام', 'role': 'instructor'}
            if any(r in role for r in ['admin', 'رئيس', 'مشرف']):
                self.instructors_data[emp_id]['role'] = 'admin'

    def create_sample_permissions(self):
        for emp_id in self.instructors_data:
            self.instructors_data[emp_id]['role'] = 'admin' if emp_id.endswith('1') else 'instructor'

    def load_schedule_data(self, filenames):
        frames = []
        for f in filenames:
            df = self.read_file(f)
            if not df.empty:
                frames.append(df)
        if frames:
            self.student_schedule_data = pd.concat(frames, ignore_index=True)
        else:
            self.create_sample_schedules()
            self.data_load_status['sample_schedules'] = True

    def create_sample_schedules(self):
        self.student_schedule_data = pd.DataFrame({
            'StudentID': ['2001', '2002'],
            'Course': ['Math', 'CS'],
            'Instructor': ['Alice', 'Bob'],
        })

    def load_students_data(self, filename):
        df = self.read_file(filename)
        if df.empty:
            self.create_sample_students()
            self.data_load_status['sample_students'] = True
        else:
            self.students_info_data = df

    def create_sample_students(self):
        self.students_info_data = pd.DataFrame({
            'ID': ['2001', '2002'],
            'Name': ['Student1', 'Student2'],
        })

    def load_instructor_schedule_data(self, filename):
        df = self.read_file(filename)
        if df.empty:
            self.create_sample_instructor_schedule()
            self.data_load_status['sample_instructor_schedule'] = True
        else:
            self.instructor_schedule_data = df

    def create_sample_instructor_schedule(self):
        self.instructor_schedule_data = pd.DataFrame({
            'EmpID': ['1001', '1002'],
            'Course': ['Math', 'CS'],
        })

    def load_saved_reports(self):
        if os.path.exists('reports.json'):
            try:
                with open('reports.json', 'r', encoding='utf-8') as f:
                    self.reports_data = json.load(f)
            except Exception:
                self.reports_data = []
        else:
            self.reports_data = []

    def save_reports(self):
        try:
            with open('reports.json', 'w', encoding='utf-8') as f:
                json.dump(self.reports_data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def create_login_interface(self):
        for w in self.root.winfo_children():
            w.destroy()
        frame = ttk.Frame(self.root)
        frame.pack(pady=20)
        ttk.Label(frame, text='نظام إدارة التدريب والحرمان', font=('Arial', 16)).pack()
        ttk.Label(frame, text='مبادرة كامل (المحسن)').pack()
        ttk.Label(frame, text=self.get_data_status_text(), foreground='blue').pack(pady=10)
        login_frame = ttk.Frame(frame)
        login_frame.pack(pady=20)
        ttk.Label(login_frame, text='رقم الموظف أو الاسم:').grid(row=0, column=0, padx=5)
        self.login_entry = ttk.Entry(login_frame)
        self.login_entry.grid(row=0, column=1, padx=5)
        ttk.Button(login_frame, text='دخول', command=self.login).grid(row=0, column=2, padx=5)
        ttk.Button(frame, text='عرض المستخدمين المتاحين', command=self.show_available_users).pack()
        if 'sample_instructors' in self.data_load_status:
            ttk.Label(frame, text='يتم استخدام بيانات تجريبية', foreground='red').pack()

    def get_data_status_text(self):
        if not self.data_integrity_issues:
            return 'جميع البيانات تم تحميلها بنجاح'
        return 'مشاكل في البيانات: ' + '; '.join(self.data_integrity_issues)

    def show_available_users(self):
        users = [f"{eid}: {info['name']} ({info['role']})" for eid, info in self.instructors_data.items()]
        messagebox.showinfo('المستخدمون', '\n'.join(users))

    def login(self):
        value = self.login_entry.get().strip()
        matches = [eid for eid, info in self.instructors_data.items() if value in (eid, info['name'])]
        if not matches:
            messagebox.showerror('خطأ', 'المستخدم غير موجود')
            return
        if len(matches) > 1:
            self.show_multiple_matches(matches)
            return
        self.current_user = matches[0]
        role = self.instructors_data[self.current_user].get('role', 'instructor')
        if role == 'admin':
            self.create_admin_interface()
        else:
            self.create_instructor_interface()

    def show_multiple_matches(self, matches):
        messagebox.showinfo('اختر', 'وجد أكثر من مستخدم: ' + ', '.join(matches))

    def show_login_help(self):
        messagebox.showinfo('مساعدة', 'أدخل رقم الموظف أو اسم المدرب')

    def create_instructor_interface(self):
        for w in self.root.winfo_children():
            w.destroy()
        nb = ttk.Notebook(self.root)
        nb.pack(fill='both', expand=True)
        self.create_course_selection_tab(nb)
        self.create_student_selection_tab(nb)
        self.create_data_entry_tab(nb)
        ttk.Button(self.root, text='تسجيل الخروج', command=self.logout).pack(pady=10)

    def create_course_selection_tab(self, nb):
        tab = ttk.Frame(nb)
        nb.add(tab, text='الدورات')
        ttk.Label(tab, text='اختر دورة').pack()
        self.course_list = tk.Listbox(tab)
        self.course_list.pack(fill='both', expand=True)
        courses = self.load_instructor_courses()
        for c in courses:
            self.course_list.insert(tk.END, c)
        self.course_list.bind('<<ListboxSelect>>', self.on_course_select)

    def load_instructor_courses(self):
        if self.instructor_schedule_data.empty:
            return []
        df = self.instructor_schedule_data
        emp_name = self.instructors_data[self.current_user]['name']
        courses = df[df.iloc[:,0] == emp_name].iloc[:,1].unique().tolist()
        return courses

    def on_course_select(self, event):
        sel = event.widget.curselection()
        if sel:
            self.selected_course = event.widget.get(sel[0])
            self.select_course()

    def select_course(self):
        messagebox.showinfo('الدورة', f'تم اختيار الدورة {self.selected_course}')

    def create_student_selection_tab(self, nb):
        tab = ttk.Frame(nb)
        nb.add(tab, text='الطلاب')
        ttk.Label(tab, text='اختر الطالب').pack()
        self.student_list = tk.Listbox(tab)
        self.student_list.pack(fill='both', expand=True)
        students = self.load_course_students()
        for s in students:
            self.student_list.insert(tk.END, s)
        self.student_list.bind('<<ListboxSelect>>', self.on_student_select)

    def load_course_students(self):
        if self.student_schedule_data.empty or not self.selected_course:
            return []
        df = self.student_schedule_data
        stus = df[df['Course'] == self.selected_course]['StudentID'].tolist()
        return stus

    def on_student_select(self, event):
        sel = event.widget.curselection()
        if sel:
            self.selected_student = event.widget.get(sel[0])
            self.select_student()

    def select_student(self):
        messagebox.showinfo('الطالب', f'تم اختيار الطالب {self.selected_student}')

    def create_data_entry_tab(self, nb):
        tab = ttk.Frame(nb)
        nb.add(tab, text='الإدخال')
        ttk.Label(tab, text='نموذج بيانات').pack()
        self.data_text = scrolledtext.ScrolledText(tab, height=10)
        self.data_text.pack(fill='both', expand=True)
        ttk.Button(tab, text='إرسال', command=self.submit_data).pack(pady=5)

    def go_back_to_courses(self):
        pass

    def go_back_to_students(self):
        pass

    def preview_data(self):
        pass

    def validate_form(self):
        return True

    def submit_data(self):
        if not self.validate_form():
            return
        report = {
            'student': self.selected_student,
            'course': self.selected_course,
            'data': self.data_text.get('1.0', tk.END),
            'status': 'pending',
            'created': datetime.now().isoformat(),
        }
        self.reports_data.append(report)
        self.save_reports()
        messagebox.showinfo('تم', 'تم إرسال التقرير')
        self.data_text.delete('1.0', tk.END)

    def export_to_excel(self):
        pass

    def reset_form(self):
        self.data_text.delete('1.0', tk.END)

    def create_admin_interface(self):
        for w in self.root.winfo_children():
            w.destroy()
        nb = ttk.Notebook(self.root)
        nb.pack(fill='both', expand=True)
        self.create_statistics_tab(nb)
        self.create_reports_tab(nb)
        ttk.Button(self.root, text='تسجيل الخروج', command=self.logout).pack(pady=10)

    def create_statistics_tab(self, nb):
        tab = ttk.Frame(nb)
        nb.add(tab, text='إحصائيات')
        ttk.Label(tab, text=f'عدد التقارير: {len(self.reports_data)}').pack()

    def create_reports_tab(self, nb):
        tab = ttk.Frame(nb)
        nb.add(tab, text='التقارير')
        self.reports_list = tk.Listbox(tab)
        self.reports_list.pack(fill='both', expand=True)
        self.load_reports_table()
        ttk.Button(tab, text='موافقة', command=self.approve_report).pack(side='left')
        ttk.Button(tab, text='رفض', command=self.reject_report).pack(side='left')

    def filter_reports(self):
        pass

    def refresh_reports(self):
        self.load_reports_table()

    def load_reports_table(self):
        self.reports_list.delete(0, tk.END)
        for idx, rep in enumerate(self.reports_data):
            self.reports_list.insert(tk.END, f"{idx}: {rep['student']} - {rep['status']}")

    def get_selected_report_id(self):
        sel = self.reports_list.curselection()
        if sel:
            return int(sel[0])
        return None

    def approve_report(self):
        rep_id = self.get_selected_report_id()
        if rep_id is not None:
            self.update_report_status(rep_id, 'approved')

    def reject_report(self):
        rep_id = self.get_selected_report_id()
        if rep_id is not None:
            self.update_report_status(rep_id, 'rejected')

    def update_report_status(self, rep_id, status):
        self.reports_data[rep_id]['status'] = status
        self.save_reports()
        self.refresh_reports()

    def show_report_details(self):
        pass

    def approve_from_details(self):
        pass

    def reject_from_details(self):
        pass

    def logout(self):
        self.current_user = None
        self.selected_course = None
        self.selected_student = None
        self.create_login_interface()

    def run(self):
        self.root.mainloop()


print('🚀 TrainingManagementSystem class generated successfully')

if __name__ == '__main__':
    try:
        import pandas
        import openpyxl
    except ImportError:
        print('pandas or openpyxl not available')

    excel_files = [f for f in os.listdir('.') if f.endswith(('.csv', '.xlsx'))]
    if len(excel_files) < 3:
        create_sample_excel_files()
    app = TrainingManagementSystem()
    app.run()
