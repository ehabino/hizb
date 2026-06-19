import React, { useState, useEffect, useMemo, useCallback } from 'react';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, Legend, ResponsiveContainer,
  PieChart, Pie, Cell
} from 'recharts';
import { 
  LayoutDashboard, FileText, CheckSquare, Users, Settings, Moon, Sun, 
  Upload, Download, Plus, Trash2, Edit, Save, Clock, AlertCircle, CheckCircle2,
  Search, Filter, ChevronDown, Calendar
} from 'lucide-react';

// --- الثوابت والإعدادات الافتراضية ---
const COLORS = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6'];
const STATUS_COLORS = {
  'تم': 'bg-emerald-100 text-emerald-800 dark:bg-emerald-900/30 dark:text-emerald-400',
  'تجديد': 'bg-amber-100 text-amber-800 dark:bg-amber-900/30 dark:text-amber-400',
  'إلغاء': 'bg-rose-100 text-rose-800 dark:bg-rose-900/30 dark:text-rose-400',
  'قيد التنفيذ': 'bg-blue-100 text-blue-800 dark:bg-blue-900/30 dark:text-blue-400'
};

const AGENDA_ITEMS = [
  "الكلمة الإفتتاحية",
  "التفقد العلمي للأعضاء (مذاكرة - حضور)",
  "تسجيل الإقتراحات",
  "متابعة التكليفات",
  "الخطة",
  "متابعة شغل الأرض",
  "أنشطة مجتمعية",
  "مناقشة الإقتراحات القديمة",
  "مناقشة الإقتراحات الجديدة"
];

const TECHNICIANS_ITEMS = [
  "الإطلاع على المحضر",
  "متابعة السكرتير",
  "الكلمة الإفتتاحية (إيمانية)",
  "التفقد العلمي للأعضاء (مذاكرة - حضور)",
  "تسجيل الإقتراحات",
  "متابعة التكليفات",
  "الخطة",
  "متابعة شغل الأرض",
  "جولات شهرية",
  "أنشطة مجتمعية",
  "مناقشة الإقتراحات القديمة",
  "مناقشة الإقتراحات الجديدة"
];

// --- المكون الرئيسي للتطبيق ---
export default function App() {
  // حالة التطبيق العامة
  const [activeTab, setActiveTab] = useState('dashboard');
  const [isDarkMode, setIsDarkMode] = useState(false);
  const [toast, setToast] = useState(null);
  const [isDirty, setIsDirty] = useState(false); // لتتبع التعديلات غير المحفوظة
  const [currentWeek, setCurrentWeek] = useState('الأسبوع الأول / شعبان - 1447');
  
  // حالة البيانات
  const [members, setMembers] = useState([{ id: 1, name: 'محمد إيهاب' }, { id: 2, name: 'يوسف محمد' }, { id: 3, name: 'آدم' }, { id: 4, name: 'أحمد طاهر' }]);
  const [weeks, setWeeks] = useState(['الأسبوع الرابع / رجب - 1447', 'الأسبوع الأول / شعبان - 1447', 'الأسبوع الثاني / شعبان - 1447', 'الأسبوع الثالث / شعبان - 1447', 'الأسبوع الرابع / شعبان - 1447']);
  const [tasks, setTasks] = useState([
    { id: 1, week: 'الأسبوع الأول / شعبان - 1447', text: 'عمل جروب لتنظيم الجولات', assignee: 'يوسف محمد', status: 'تم', created: '2023-10-01', modified: '2023-10-01' },
    { id: 2, week: 'الأسبوع الأول / شعبان - 1447', text: 'متابعة نزول الجولات', assignee: 'محمد إيهاب', status: 'قيد التنفيذ', created: '2023-10-02', modified: '2023-10-02' },
    { id: 3, week: 'الأسبوع الثاني / شعبان - 1447', text: 'ترتيب فاعلية جديدة', assignee: 'أحمد طاهر', status: 'إلغاء', created: '2023-10-05', modified: '2023-10-06' },
  ]);
  const [minutes, setMinutes] = useState({}); // { 'week1': { attendance: {}, items: {} } }
  const [techItems, setTechItems] = useState({}); // { 'week1': { 'item1': true } }

  // --- دوال مساعدة ---
  const showToast = (msg, type = 'success') => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 3000);
  };

  const markDirty = () => setIsDirty(true);

  // تحميل مكتبة Excel ديناميكياً 
  useEffect(() => {
    if (typeof window !== 'undefined' && !window.XLSX) {
      const script = document.createElement('script');
      script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
      script.async = true;
      document.body.appendChild(script);
    }
  }, []);

  // منع فقدان البيانات عند التحديث
  useEffect(() => {
    const handleBeforeUnload = (e) => {
      if (isDirty) {
        e.preventDefault();
        e.returnValue = '';
      }
    };
    window.addEventListener('beforeunload', handleBeforeUnload);
    return () => window.removeEventListener('beforeunload', handleBeforeUnload);
  }, [isDirty]);

  // تبديل الوضع الليلي
  useEffect(() => {
    if (isDarkMode) document.documentElement.classList.add('dark');
    else document.documentElement.classList.remove('dark');
  }, [isDarkMode]);

  // --- معالجة ملف Excel ---
  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    if (!window.XLSX) {
      showToast('جاري إعداد مكتبة الملفات، يرجى المحاولة بعد لحظات', 'error');
      return;
    }

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target.result;
        const wb = window.XLSX.read(bstr, { type: 'binary' });
        
        // استخراج التكليفات كمثال (يفترض وجود شيت باسم "تكليفات" أو رقم 2)
        const taskSheetName = wb.SheetNames.find(n => n.includes('تكليف')) || wb.SheetNames[2];
        if (taskSheetName) {
          const taskData = window.XLSX.utils.sheet_to_json(wb.Sheets[taskSheetName], { header: 1 });
          const parsedTasks = [];
          const parsedMembers = new Set();
          
          taskData.forEach((row, idx) => {
            if (idx > 1 && row[1] && row[2]) { // تخطي الرؤوس
              parsedTasks.push({
                id: Date.now() + idx,
                week: row[1] || currentWeek,
                text: row[2] || '',
                assignee: row[3] || 'غير محدد',
                status: row[4] || 'قيد التنفيذ',
                created: new Date().toISOString().split('T')[0],
                modified: new Date().toISOString().split('T')[0]
              });
              if(row[3]) parsedMembers.add(row[3]);
            }
          });
          
          if(parsedTasks.length > 0) setTasks(parsedTasks);
          if(parsedMembers.size > 0) {
            setMembers(Array.from(parsedMembers).map((m, i) => ({ id: i+1, name: m })));
          }
        }
        showToast('تم تحميل البيانات من Excel بنجاح');
        setIsDirty(false);
      } catch (error) {
        showToast('حدث خطأ أثناء قراءة الملف', 'error');
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleExportExcel = () => {
    if (!window.XLSX) {
      showToast('جاري إعداد مكتبة الملفات، يرجى المحاولة بعد لحظات', 'error');
      return;
    }

    // بناء ملف Excel من البيانات الحالية
    const wb = window.XLSX.utils.book_new();
    
    // شيت التكليفات
    const taskWsData = [['التكليفات'], ['الأسبوع', 'التكليف', 'مسئول التنفيذ', 'حالة التنفيذ', 'تاريخ الإنشاء', 'تاريخ التعديل']];
    tasks.forEach(t => taskWsData.push([t.week, t.text, t.assignee, t.status, t.created, t.modified]));
    const wsTasks = window.XLSX.utils.aoa_to_sheet(taskWsData);
    window.XLSX.utils.book_append_sheet(wb, wsTasks, "التكليفات");

    // تحميل الملف
    window.XLSX.writeFile(wb, "hizb_updated.xlsx");
    showToast('تم تصدير وحفظ الملف بنجاح');
    setIsDirty(false);
  };

  // --- الإحصائيات (مدمجة) ---
  const stats = useMemo(() => {
    const totalTasks = tasks.length;
    const completedTasks = tasks.filter(t => t.status === 'تم').length;
    const completionRate = totalTasks ? Math.round((completedTasks / totalTasks) * 100) : 0;
    
    // أداء الأعضاء
    const memberStats = {};
    members.forEach(m => memberStats[m.name] = { name: m.name, total: 0, completed: 0 });
    tasks.forEach(t => {
      if (memberStats[t.assignee]) {
        memberStats[t.assignee].total++;
        if (t.status === 'تم') memberStats[t.assignee].completed++;
      }
    });
    
    const chartData = Object.values(memberStats).map(m => ({
      name: m.name,
      'نسبة الإنجاز': m.total ? Math.round((m.completed / m.total) * 100) : 0,
      'الإجمالي': m.total
    })).sort((a, b) => b['نسبة الإنجاز'] - a['نسبة الإنجاز']);

    // حالة التكليفات
    const statusCount = { 'تم': 0, 'قيد التنفيذ': 0, 'تجديد': 0, 'إلغاء': 0 };
    tasks.forEach(t => { if(statusCount[t.status] !== undefined) statusCount[t.status]++; });
    const pieData = Object.keys(statusCount).map(k => ({ name: k, value: statusCount[k] })).filter(d => d.value > 0);

    return { totalTasks, completedTasks, completionRate, chartData, pieData };
  }, [tasks, members]);


  // --- المكونات الفرعية (مدمجة لتكون في ملف واحد) ---

  // 1. لوحة التحكم
  const Dashboard = () => (
    <div className="space-y-6 animate-fade-in">
      <h2 className="text-2xl font-bold text-gray-800 dark:text-white mb-6">ملخص الأداء والتحليلات</h2>
      
      {/* Cards */}
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
        <StatCard icon={<CheckSquare size={24} />} title="إجمالي التكليفات" value={stats.totalTasks} color="bg-blue-500" />
        <StatCard icon={<CheckCircle2 size={24} />} title="نسبة الإنجاز الكلية" value={`${stats.completionRate}%`} color="bg-emerald-500" />
        <StatCard icon={<Users size={24} />} title="عدد الأعضاء" value={members.length} color="bg-purple-500" />
        <StatCard icon={<Calendar size={24} />} title="الأسبوع الحالي" value={currentWeek.split('/')[0]} subtitle={currentWeek.split('/')[1]} color="bg-amber-500" />
      </div>

      {/* Charts */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mt-8">
        <div className="bg-white dark:bg-slate-800 p-6 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-700">
          <h3 className="text-lg font-semibold mb-6 text-slate-800 dark:text-white">نسبة إنجاز الأعضاء</h3>
          <div className="h-72">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={stats.chartData} layout="vertical" margin={{ top: 5, right: 30, left: 20, bottom: 5 }}>
                <CartesianGrid strokeDasharray="3 3" opacity={0.2} />
                <XAxis type="number" domain={[0, 100]} />
                <YAxis dataKey="name" type="category" width={100} tick={{fill: isDarkMode ? '#e2e8f0' : '#475569'}} />
                <RechartsTooltip cursor={{fill: 'transparent'}} contentStyle={{borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)'}} />
                <Bar dataKey="نسبة الإنجاز" fill="#3b82f6" radius={[0, 4, 4, 0]} barSize={20} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="bg-white dark:bg-slate-800 p-6 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-700">
          <h3 className="text-lg font-semibold mb-6 text-slate-800 dark:text-white">حالة التكليفات</h3>
          <div className="h-72">
            <ResponsiveContainer width="100%" height="100%">
              <PieChart>
                <Pie data={stats.pieData} cx="50%" cy="50%" innerRadius={60} outerRadius={80} paddingAngle={5} dataKey="value">
                  {stats.pieData.map((entry, index) => (
                    <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                  ))}
                </Pie>
                <RechartsTooltip contentStyle={{borderRadius: '8px', border: 'none'}} />
                <Legend verticalAlign="bottom" height={36} />
              </PieChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>
    </div>
  );

  // 2. محضر الجلسة
  const MinutesTab = () => {
    const weekData = minutes[currentWeek] || { attendance: {}, items: {} };

    const handleAttendance = (memberId, isChecked) => {
      const time = isChecked ? new Date().toLocaleTimeString('ar-EG', { hour: '2-digit', minute:'2-digit', second:'2-digit' }) : null;
      setMinutes(prev => ({
        ...prev,
        [currentWeek]: {
          ...prev[currentWeek],
          attendance: { ...prev[currentWeek]?.attendance, [memberId]: time }
        }
      }));
      markDirty();
    };

    const handleItemChange = (itemIdx, field, value) => {
      setMinutes(prev => ({
        ...prev,
        [currentWeek]: {
          ...prev[currentWeek],
          items: {
            ...prev[currentWeek]?.items,
            [itemIdx]: { ...prev[currentWeek]?.items?.[itemIdx], [field]: value }
          }
        }
      }));
      markDirty();
    };

    return (
      <div className="space-y-8 animate-fade-in">
        {/* Header Setup */}
        <div className="bg-white dark:bg-slate-800 p-6 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-700 flex flex-wrap gap-4 items-center justify-between">
          <div>
            <h2 className="text-xl font-bold text-slate-800 dark:text-white mb-2">تسجيل محضر الجلسة</h2>
            <p className="text-sm text-slate-500 dark:text-slate-400">حدد الأسبوع وسجل الحضور والبنود تلقائياً</p>
          </div>
          <div className="w-64">
            <label className="block text-sm font-medium mb-1 text-slate-700 dark:text-slate-300">الأسبوع الهجري</label>
            <select 
              value={currentWeek} 
              onChange={(e) => setCurrentWeek(e.target.value)}
              className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 rounded-lg px-4 py-2 text-slate-800 dark:text-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
            >
              {weeks.map(w => <option key={w} value={w}>{w}</option>)}
            </select>
          </div>
        </div>

        {/* Attendance */}
        <div className="bg-white dark:bg-slate-800 p-6 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-700">
          <h3 className="text-lg font-semibold mb-4 text-slate-800 dark:text-white flex items-center gap-2">
            <Users className="text-blue-500" size={20} /> تسجيل الحضور
          </h3>
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
            {members.map(m => (
              <label key={m.id} className="flex items-center justify-between p-4 rounded-xl border border-slate-100 dark:border-slate-700 hover:bg-slate-50 dark:hover:bg-slate-750 cursor-pointer transition-colors">
                <div className="flex items-center gap-3">
                  <input 
                    type="checkbox" 
                    className="w-5 h-5 rounded border-slate-300 text-blue-600 focus:ring-blue-500 dark:border-slate-600 dark:bg-slate-700"
                    checked={!!weekData.attendance?.[m.id]}
                    onChange={(e) => handleAttendance(m.id, e.target.checked)}
                  />
                  <span className="font-medium text-slate-800 dark:text-slate-200">{m.name}</span>
                </div>
                {weekData.attendance?.[m.id] && (
                  <span className="text-xs bg-blue-100 text-blue-800 dark:bg-blue-900/50 dark:text-blue-300 py-1 px-2 rounded-md font-mono flex items-center gap-1">
                    <Clock size={12} /> {weekData.attendance[m.id]}
                  </span>
                )}
              </label>
            ))}
          </div>
        </div>

        {/* Agenda Items */}
        <div className="bg-white dark:bg-slate-800 p-6 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-700">
          <h3 className="text-lg font-semibold mb-6 text-slate-800 dark:text-white flex items-center gap-2">
            <FileText className="text-purple-500" size={20} /> بنود المحضر الأسبوعي
          </h3>
          <div className="space-y-6">
            {AGENDA_ITEMS.map((item, idx) => {
              const itemData = weekData.items?.[idx] || { status: '', notes: '' };
              return (
                <div key={idx} className="p-5 rounded-xl border border-slate-100 dark:border-slate-700 bg-slate-50/50 dark:bg-slate-800/50">
                  <div className="flex flex-col md:flex-row md:items-center justify-between mb-4 gap-4">
                    <h4 className="font-semibold text-slate-800 dark:text-slate-200 text-lg">{item}</h4>
                    <div className="flex gap-4">
                      <label className="flex items-center gap-2 cursor-pointer">
                        <input type="radio" name={`status-${idx}`} value="تم" checked={itemData.status === 'تم'} onChange={() => handleItemChange(idx, 'status', 'تم')} className="w-4 h-4 text-emerald-600" />
                        <span className="text-sm font-medium text-slate-700 dark:text-slate-300">تم</span>
                      </label>
                      <label className="flex items-center gap-2 cursor-pointer">
                        <input type="radio" name={`status-${idx}`} value="لم تتم" checked={itemData.status === 'لم تتم'} onChange={() => handleItemChange(idx, 'status', 'لم تتم')} className="w-4 h-4 text-rose-600" />
                        <span className="text-sm font-medium text-slate-700 dark:text-slate-300">لم تتم</span>
                      </label>
                    </div>
                  </div>
                  <textarea 
                    placeholder="أضف ملاحظاتك هنا..."
                    value={itemData.notes}
                    onChange={(e) => handleItemChange(idx, 'notes', e.target.value)}
                    className="w-full bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-700 rounded-lg p-3 text-sm text-slate-800 dark:text-slate-200 focus:ring-2 focus:ring-blue-500 outline-none resize-none h-20 transition-all"
                  ></textarea>
                </div>
              );
            })}
          </div>
        </div>
      </div>
    );
  };

  // 3. التكليفات
  const TasksTab = () => {
    const [searchTerm, setSearchTerm] = useState('');
    const [filterStatus, setFilterStatus] = useState('الكل');

    const filteredTasks = tasks.filter(t => 
      (t.text.includes(searchTerm) || t.assignee.includes(searchTerm)) &&
      (filterStatus === 'الكل' || t.status === filterStatus)
    );

    const addTask = () => {
      const newTask = {
        id: Date.now(),
        week: currentWeek,
        text: 'تكليف جديد...',
        assignee: members[0]?.name || 'غير محدد',
        status: 'قيد التنفيذ',
        created: new Date().toISOString().split('T')[0],
        modified: new Date().toISOString().split('T')[0]
      };
      setTasks([newTask, ...tasks]);
      markDirty();
    };

    const updateTask = (id, field, value) => {
      setTasks(tasks.map(t => t.id === id ? { ...t, [field]: value, modified: new Date().toISOString().split('T')[0] } : t));
      markDirty();
    };

    const deleteTask = (id) => {
      if(window.confirm('هل أنت متأكد من حذف هذا التكليف؟')) {
        setTasks(tasks.filter(t => t.id !== id));
        markDirty();
      }
    };

    return (
      <div className="space-y-6 animate-fade-in h-full flex flex-col">
        <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 bg-white dark:bg-slate-800 p-6 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-700">
          <div>
            <h2 className="text-xl font-bold text-slate-800 dark:text-white">إدارة التكليفات</h2>
            <p className="text-sm text-slate-500 dark:text-slate-400">تتبع وأدر مهام الأعضاء بكفاءة</p>
          </div>
          <button onClick={addTask} className="bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-700 hover:to-indigo-700 text-white px-5 py-2.5 rounded-xl font-medium shadow-md shadow-blue-500/30 flex items-center gap-2 transition-all transform hover:scale-105 active:scale-95">
            <Plus size={18} /> إضافة تكليف
          </button>
        </div>

        <div className="flex flex-col md:flex-row gap-4">
          <div className="relative flex-1">
            <Search className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400" size={18} />
            <input 
              type="text" 
              placeholder="ابحث في التكليفات..." 
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="w-full bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-xl py-2.5 pr-10 pl-4 text-slate-800 dark:text-slate-200 focus:ring-2 focus:ring-blue-500 outline-none shadow-sm"
            />
          </div>
          <div className="relative w-full md:w-48">
            <Filter className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400" size={18} />
            <select 
              value={filterStatus}
              onChange={(e) => setFilterStatus(e.target.value)}
              className="w-full bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-xl py-2.5 pr-10 pl-4 text-slate-800 dark:text-slate-200 focus:ring-2 focus:ring-blue-500 outline-none shadow-sm appearance-none"
            >
              <option value="الكل">جميع الحالات</option>
              <option value="تم">تم</option>
              <option value="قيد التنفيذ">قيد التنفيذ</option>
              <option value="تجديد">تجديد</option>
              <option value="إلغاء">إلغاء</option>
            </select>
          </div>
        </div>

        <div className="bg-white dark:bg-slate-800 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-700 overflow-hidden flex-1">
          <div className="overflow-x-auto">
            <table className="w-full text-right border-collapse">
              <thead>
                <tr className="bg-slate-50 dark:bg-slate-900 border-b border-slate-200 dark:border-slate-700">
                  <th className="p-4 font-semibold text-slate-600 dark:text-slate-300 w-1/4">التكليف</th>
                  <th className="p-4 font-semibold text-slate-600 dark:text-slate-300">المسئول</th>
                  <th className="p-4 font-semibold text-slate-600 dark:text-slate-300">الأسبوع</th>
                  <th className="p-4 font-semibold text-slate-600 dark:text-slate-300">الحالة</th>
                  <th className="p-4 font-semibold text-slate-600 dark:text-slate-300 text-center">إجراء</th>
                </tr>
              </thead>
              <tbody>
                {filteredTasks.length === 0 ? (
                  <tr><td colSpan="5" className="p-8 text-center text-slate-500">لا توجد تكليفات مطابقة</td></tr>
                ) : filteredTasks.map(t => (
                  <tr key={t.id} className="border-b border-slate-100 dark:border-slate-700/50 hover:bg-slate-50/50 dark:hover:bg-slate-800/50 transition-colors group">
                    <td className="p-4">
                      <input 
                        type="text" value={t.text} onChange={(e) => updateTask(t.id, 'text', e.target.value)}
                        className="w-full bg-transparent border-none focus:ring-1 focus:ring-blue-500 rounded px-2 py-1 outline-none text-slate-800 dark:text-slate-200"
                      />
                    </td>
                    <td className="p-4">
                      <select value={t.assignee} onChange={(e) => updateTask(t.id, 'assignee', e.target.value)} className="bg-transparent border-none focus:ring-1 focus:ring-blue-500 rounded outline-none text-slate-800 dark:text-slate-200 cursor-pointer">
                        {members.map(m => <option key={m.id} value={m.name}>{m.name}</option>)}
                      </select>
                    </td>
                    <td className="p-4 text-sm text-slate-600 dark:text-slate-400">{t.week.split('/')[0]}</td>
                    <td className="p-4">
                      <select 
                        value={t.status} onChange={(e) => updateTask(t.id, 'status', e.target.value)}
                        className={`text-xs font-bold px-3 py-1.5 rounded-full outline-none cursor-pointer appearance-none border-none ${STATUS_COLORS[t.status] || STATUS_COLORS['قيد التنفيذ']}`}
                      >
                        <option value="تم">تم</option>
                        <option value="قيد التنفيذ">قيد التنفيذ</option>
                        <option value="تجديد">تجديد</option>
                        <option value="إلغاء">إلغاء</option>
                      </select>
                    </td>
                    <td className="p-4 text-center">
                      <button onClick={() => deleteTask(t.id)} className="text-slate-400 hover:text-rose-500 transition-colors opacity-0 group-hover:opacity-100">
                        <Trash2 size={18} />
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    );
  };

  // 4. بنود الفنيين
  const TechItemsTab = () => {
    const currentTech = techItems[currentWeek] || {};
    
    const toggleItem = (item) => {
      setTechItems(prev => ({
        ...prev, [currentWeek]: { ...prev[currentWeek], [item]: !prev[currentWeek]?.[item] }
      }));
      markDirty();
    };

    return (
      <div className="max-w-4xl mx-auto space-y-6 animate-fade-in">
         <div className="bg-white dark:bg-slate-800 p-6 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-700">
            <h2 className="text-xl font-bold text-slate-800 dark:text-white mb-2">بنود الفنيين - {currentWeek}</h2>
            <p className="text-sm text-slate-500 dark:text-slate-400 mb-6">قائمة التحقق الخاصة بالفنيين للأسبوع المحدد.</p>
            
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              {TECHNICIANS_ITEMS.map((item, idx) => {
                const isDone = !!currentTech[item];
                return (
                  <div key={idx} onClick={() => toggleItem(item)} className={`p-4 rounded-xl border cursor-pointer transition-all flex items-center justify-between ${isDone ? 'bg-emerald-50 border-emerald-200 dark:bg-emerald-900/20 dark:border-emerald-800' : 'bg-slate-50 border-slate-200 dark:bg-slate-800 dark:border-slate-700 hover:border-blue-300 dark:hover:border-blue-700'}`}>
                    <span className={`font-medium ${isDone ? 'text-emerald-800 dark:text-emerald-300' : 'text-slate-700 dark:text-slate-300'}`}>{item}</span>
                    <div className={`w-6 h-6 rounded-full flex items-center justify-center border-2 ${isDone ? 'bg-emerald-500 border-emerald-500 text-white' : 'border-slate-300 dark:border-slate-600'}`}>
                      {isDone && <CheckCircle2 size={16} />}
                    </div>
                  </div>
                );
              })}
            </div>
         </div>
      </div>
    );
  }

  // --- واجهة التطبيق الهيكلية ---
  const TABS = [
    { id: 'dashboard', label: 'الرئيسية', icon: <LayoutDashboard size={20} /> },
    { id: 'minutes', label: 'محضر الجلسة', icon: <FileText size={20} /> },
    { id: 'tasks', label: 'التكليفات', icon: <CheckSquare size={20} /> },
    { id: 'tech', label: 'بنود الفنيين', icon: <Settings size={20} /> }
  ];

  return (
    <div dir="rtl" className="flex h-screen bg-slate-50 dark:bg-slate-900 font-sans text-slate-900 dark:text-slate-100 overflow-hidden transition-colors duration-300">
      
      {/* Sidebar */}
      <aside className="w-64 bg-white dark:bg-slate-950 border-l border-slate-200 dark:border-slate-800 flex flex-col shadow-lg z-10 hidden md:flex">
        <div className="p-6 border-b border-slate-100 dark:border-slate-800 flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-blue-600 to-indigo-600 flex items-center justify-center text-white font-bold text-xl shadow-lg shadow-blue-500/30">ح</div>
          <div>
            <h1 className="font-bold text-lg text-slate-800 dark:text-white leading-tight">نظام اللجان</h1>
            <span className="text-xs text-slate-500 dark:text-slate-400">الإصدار 2.0</span>
          </div>
        </div>
        
        <nav className="flex-1 px-4 py-6 space-y-2 overflow-y-auto">
          {TABS.map(tab => (
            <button
              key={tab.id}
              onClick={() => setActiveTab(tab.id)}
              className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all font-medium ${activeTab === tab.id ? 'bg-blue-50 text-blue-700 dark:bg-blue-900/30 dark:text-blue-400' : 'text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-800'}`}
            >
              {tab.icon} {tab.label}
            </button>
          ))}
        </nav>

        <div className="p-4 border-t border-slate-100 dark:border-slate-800 space-y-3">
           <label className="w-full flex items-center justify-center gap-2 px-4 py-2.5 rounded-xl border border-dashed border-slate-300 dark:border-slate-700 text-sm font-medium text-slate-600 dark:text-slate-400 hover:bg-slate-50 dark:hover:bg-slate-800 cursor-pointer transition-colors">
              <Upload size={16} /> رفع Excel
              <input type="file" accept=".xlsx, .xls, .csv" className="hidden" onChange={handleFileUpload} />
           </label>
           
           <button onClick={handleExportExcel} className={`w-full flex items-center justify-center gap-2 px-4 py-2.5 rounded-xl text-sm font-medium transition-all text-white shadow-md ${isDirty ? 'bg-gradient-to-r from-amber-500 to-orange-500 hover:from-amber-600 hover:to-orange-600 shadow-amber-500/30 animate-pulse' : 'bg-gradient-to-r from-slate-700 to-slate-800 hover:from-slate-800 hover:to-slate-900 dark:from-slate-600 dark:to-slate-700 shadow-slate-500/30'}`}>
              <Save size={16} /> {isDirty ? 'حفظ التغييرات' : 'تصدير Excel'}
           </button>
        </div>
      </aside>

      {/* Main Content */}
      <main className="flex-1 flex flex-col h-full overflow-hidden relative">
        {/* Header (Mobile & Top actions) */}
        <header className="h-16 bg-white/80 dark:bg-slate-950/80 backdrop-blur-md border-b border-slate-200 dark:border-slate-800 flex items-center justify-between px-6 z-10">
          <div className="flex md:hidden items-center gap-2">
            <div className="w-8 h-8 rounded-lg bg-gradient-to-br from-blue-600 to-indigo-600 flex items-center justify-center text-white font-bold">ح</div>
            <span className="font-bold">نظام اللجان</span>
          </div>
          <div className="hidden md:flex items-center gap-2">
             {isDirty && <span className="flex items-center gap-1 text-xs font-medium text-amber-600 dark:text-amber-400 bg-amber-50 dark:bg-amber-900/30 px-3 py-1 rounded-full"><AlertCircle size={14}/> يوجد تعديلات غير محفوظة</span>}
          </div>
          
          <div className="flex items-center gap-4">
             <div className="text-sm font-medium text-slate-500 hidden sm:block">مرحباً، <span className="text-slate-800 dark:text-white">المسؤول</span></div>
             <button onClick={() => setIsDarkMode(!isDarkMode)} className="p-2 rounded-full hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors text-slate-600 dark:text-slate-300">
               {isDarkMode ? <Sun size={20} /> : <Moon size={20} />}
             </button>
          </div>
        </header>

        {/* Scrollable Content Area */}
        <div className="flex-1 overflow-y-auto p-4 md:p-8">
           {activeTab === 'dashboard' && <Dashboard />}
           {activeTab === 'minutes' && <MinutesTab />}
           {activeTab === 'tasks' && <TasksTab />}
           {activeTab === 'tech' && <TechItemsTab />}
        </div>

        {/* Toast Notification */}
        {toast && (
          <div className={`absolute bottom-6 left-6 flex items-center gap-3 px-6 py-4 rounded-xl shadow-2xl animate-slide-up z-50 text-white font-medium ${toast.type === 'error' ? 'bg-rose-500' : 'bg-slate-800 dark:bg-slate-700'}`}>
            {toast.type === 'success' ? <CheckCircle2 size={20} className="text-emerald-400" /> : <AlertCircle size={20} />}
            {toast.msg}
          </div>
        )}
      </main>

      {/* Tailwind Custom Utilities (Injected dynamically for single-file constraints) */}
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;800&display=swap');
        body { font-family: 'Tajawal', sans-serif; }
        .animate-fade-in { animation: fadeIn 0.4s ease-out forwards; }
        .animate-slide-up { animation: slideUp 0.3s cubic-bezier(0.16, 1, 0.3, 1) forwards; }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }
        @keyframes slideUp { from { opacity: 0; transform: translateY(20px); } to { opacity: 1; transform: translateY(0); } }
        /* Scrollbar Styling */
        ::-webkit-scrollbar { width: 6px; height: 6px; }
        ::-webkit-scrollbar-track { background: transparent; }
        ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 10px; }
        .dark ::-webkit-scrollbar-thumb { background: #475569; }
      `}</style>
    </div>
  );
}

// مكون فرعي للبطاقات الإحصائية
const StatCard = ({ icon, title, value, subtitle, color }) => (
  <div className="bg-white dark:bg-slate-800 p-6 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-700 flex items-center gap-4 group hover:shadow-md transition-shadow">
    <div className={`w-14 h-14 rounded-2xl flex items-center justify-center text-white shadow-lg ${color} bg-opacity-90 group-hover:scale-110 transition-transform`}>
      {icon}
    </div>
    <div>
      <h3 className="text-sm font-medium text-slate-500 dark:text-slate-400 mb-1">{title}</h3>
      <div className="text-2xl font-bold text-slate-800 dark:text-white">{value}</div>
      {subtitle && <div className="text-xs text-slate-400 mt-1">{subtitle}</div>}
    </div>
  </div>
);