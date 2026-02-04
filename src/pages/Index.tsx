import { useState, useCallback, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { ThemeToggle } from '@/components/ThemeToggle';
import { FileUpload } from '@/components/FileUpload';
import { StatsCard } from '@/components/StatsCard';
import { FilterButtons } from '@/components/FilterButtons';
import { PhoneTable, PhoneEntry } from '@/components/PhoneTable';
import { Loader2 } from 'lucide-react';

type FilterType = 'all' | 'incomplete' | 'complete';

interface ExcelRow {
  'رقم الجوال'?: string | number;
  'شعار'?: string;
  'الشهادة'?: string;
  'المنتجات'?: string;
  'اللوكيشن'?: string;
  'حساب بنكي'?: string;
}

function generateCustomMessage(row: ExcelRow): string | null {
  const observations: string[] = [];

  if (row['شعار'] === 'لا') {
    observations.push('- لم يتم رفع الشعار، هل يوجد لديكم شعار؟');
  }
  if (row['الشهادة'] === 'لا') {
    observations.push('- لم يتم رفع شهادة العمل الحر أو الأسر المنتجة');
  }
  if (row['المنتجات'] === 'لا') {
    observations.push('- لم يتم إضافة ٥ منتجات');
  }
  if (row['اللوكيشن'] === 'لا') {
    observations.push('- لم يتم إضافة اللوكيشن');
  }
  if (row['حساب بنكي'] === 'لا') {
    observations.push('- لم يتم إضافة حساب بنكي');
  }

  if (observations.length === 0) {
    return null;
  }

  let message = 'السلام عليكم ورحمة الله وبركاته\n\n';
  message += 'نشكر إنضمامكم لمنصة من حولكم، ونفيدكم أنه تم مراجعة طلبكم، ووجدنا بعض الملاحظات يرجى إكمالها بأسرع وقت ممكن:\n\n';
  message += observations.join('\n\n');
  message += '\n\nهل فيه شيء معين واجهك بإكمال متجرك ؟ ممكن نساعدك فيه';

  return message;
}

function createWhatsAppLink(phone: string | number, message: string): string {
  let phoneStr = phone.toString();
  if (!phoneStr.startsWith('966')) {
    phoneStr = '966' + phoneStr;
  }
  return `https://wa.me/${phoneStr}?text=${encodeURIComponent(message)}`;
}

const Index = () => {
  const [fileName, setFileName] = useState('');
  const [loading, setLoading] = useState(false);
  const [entries, setEntries] = useState<PhoneEntry[]>([]);
  const [activeFilter, setActiveFilter] = useState<FilterType>('all');
  const [sentToBottom, setSentToBottom] = useState(true);

  const handleFileSelect = useCallback((file: File) => {
    setFileName(file.name);
    setLoading(true);

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json<ExcelRow>(firstSheet);

        const results: PhoneEntry[] = [];

        jsonData.forEach((row) => {
          const phone = row['رقم الجوال'];
          if (!phone) return;

          const message = generateCustomMessage(row);

          if (message) {
            results.push({
              phone: phone.toString(),
              link: createWhatsAppLink(phone, message),
              hasMessage: true,
              sent: false,
            });
          } else {
            const thankYouMessage = 'السلام عليكم ورحمة الله وبركاته\n\nشكراً لإكمال جميع البيانات في منصة من حولكم ✓\nنتمنى لكم التوفيق';
            results.push({
              phone: phone.toString(),
              link: createWhatsAppLink(phone, thankYouMessage),
              hasMessage: false,
              sent: false,
            });
          }
        });

        setEntries(results);
        setLoading(false);
      } catch (error) {
        console.error(error);
        alert('حدث خطأ في قراءة الملف. تأكد من أن الملف بصيغة Excel صحيحة.');
        setLoading(false);
      }
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const handleMarkAsSent = useCallback((displayIndex: number) => {
    setEntries((prev) => {
      const filtered = getFilteredEntries(prev, activeFilter, sentToBottom);
      const actualEntry = filtered[displayIndex];

      return prev.map((entry) =>
        entry === actualEntry ? { ...entry, sent: true } : entry
      );
    });
  }, [activeFilter, sentToBottom]);

  const handleAutoSendIncomplete = useCallback(() => {
    const incompleteUnsent = entries.filter((e) => e.hasMessage && !e.sent);

    if (incompleteUnsent.length === 0) return;

    // Open each WhatsApp link with a delay to avoid browser blocking
    incompleteUnsent.forEach((entry, index) => {
      setTimeout(() => {
        window.open(entry.link, '_blank');
      }, index * 1500); // 1.5 second delay between each
    });

    // Mark all as sent
    setEntries((prev) =>
      prev.map((entry) =>
        entry.hasMessage && !entry.sent ? { ...entry, sent: true } : entry
      )
    );
  }, [entries]);

  const getFilteredEntries = (data: PhoneEntry[], filter: FilterType, sortSentToBottom: boolean) => {
    let filtered = [...data];

    if (filter === 'incomplete') {
      filtered = filtered.filter((e) => e.hasMessage);
    } else if (filter === 'complete') {
      filtered = filtered.filter((e) => !e.hasMessage);
    }

    filtered.sort((a, b) => {
      const aSent = a.sent || false;
      const bSent = b.sent || false;
      if (sortSentToBottom) {
        return aSent === bSent ? 0 : aSent ? 1 : -1;
      } else {
        return aSent === bSent ? 0 : aSent ? -1 : 1;
      }
    });

    return filtered;
  };

  const filteredEntries = useMemo(
    () => getFilteredEntries(entries, activeFilter, sentToBottom),
    [entries, activeFilter, sentToBottom]
  );

  const stats = useMemo(() => {
    const total = entries.length;
    const sent = entries.filter((e) => e.sent).length;
    const remaining = total - sent;
    const incompleteTotal = entries.filter((e) => e.hasMessage).length;
    const incompleteSent = entries.filter((e) => e.hasMessage && e.sent).length;
    const incompleteRemaining = incompleteTotal - incompleteSent;
    const complete = entries.filter((e) => !e.hasMessage).length;

    return {
      total,
      sent,
      remaining,
      incompleteTotal,
      incompleteSent,
      incompleteRemaining,
      complete,
    };
  }, [entries]);

  const filterCounts = useMemo(
    () => ({
      all: entries.length,
      incomplete: entries.filter((e) => e.hasMessage).length,
      complete: entries.filter((e) => !e.hasMessage).length,
    }),
    [entries]
  );

  return (
    <div className="min-h-screen p-5">
      <div className="max-w-4xl mx-auto">
        {/* Header */}
        <div className="bg-card rounded-xl shadow-sm p-6 mb-6 relative">
          <ThemeToggle />
          <h1 className="text-2xl font-bold text-foreground mb-2">
            منصة من حولكم
          </h1>
          <p className="text-muted-foreground text-sm mb-5">
            نظام إرسال رسائل المتابعة
          </p>
          <FileUpload fileName={fileName} onFileSelect={handleFileSelect} />
        </div>

        {/* Loading */}
        {loading && (
          <div className="text-center py-10 text-primary font-medium flex items-center justify-center gap-2">
            <Loader2 className="h-5 w-5 animate-spin" />
            جاري معالجة البيانات...
          </div>
        )}

        {/* Stats */}
        {entries.length > 0 && !loading && (
          <>
            <StatsCard
              total={stats.total}
              sent={stats.sent}
              remaining={stats.remaining}
              incompleteTotal={stats.incompleteTotal}
              incompleteSent={stats.incompleteSent}
              incompleteRemaining={stats.incompleteRemaining}
              complete={stats.complete}
            />

            <FilterButtons
              activeFilter={activeFilter}
              onFilterChange={setActiveFilter}
              counts={filterCounts}
              sentToBottom={sentToBottom}
              onToggleSentPosition={() => setSentToBottom(!sentToBottom)}
              onAutoSendIncomplete={handleAutoSendIncomplete}
              incompleteRemaining={stats.incompleteRemaining}
            />

            <PhoneTable
              entries={filteredEntries}
              onMarkAsSent={handleMarkAsSent}
            />
          </>
        )}

        {/* Footer */}
        <div className="text-center text-muted-foreground mt-8 py-5 text-sm">
          منصة من حولكم © 2026
        </div>
      </div>
    </div>
  );
};

export default Index;
