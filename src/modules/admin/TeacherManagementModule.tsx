import React, { useState, useEffect } from 'react';
import {
  Users,
  UserPlus,
  FileSpreadsheet,
  Download,
  Search,
  Edit2,
  Trash2,
  Key,
  Copy,
  Check,
  RefreshCw,
  Eye,
  EyeOff,
  Briefcase,
  Mail,
  Phone,
  AlertCircle,
  ShieldCheck,
  Sparkles,
  X,
} from 'lucide-react';
import { User } from '../../types/index.ts';
import * as api from '../../services/api.ts';
import { ExcelUploadModal } from '../../components/ExcelUploadModal.tsx';
import { downloadTeacherTemplateExcel } from '../../utils/teacherTemplate.ts';

// Client-side random password generator
const generateRandomPassword = (length = 9) => {
  const upper = 'ABCDEFGHJKLMNPQRSTUVWXYZ';
  const lower = 'abcdefghjkmnpqrstuvwxyz';
  const numbers = '23456789';
  const chars = upper + lower + numbers;
  let result = '';
  result += upper[Math.floor(Math.random() * upper.length)];
  result += lower[Math.floor(Math.random() * lower.length)];
  result += numbers[Math.floor(Math.random() * numbers.length)];
  for (let i = 3; i < length; i++) {
    result += chars[Math.floor(Math.random() * chars.length)];
  }
  return result;
};

// Helper to safely parse comma-separated contacts (handles multiple phone or email strings)
export const parseCommaList = (val?: string): string[] => {
  if (!val) return [];
  return val
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean);
};

export const TeacherManagementModule: React.FC = () => {
  const [teachers, setTeachers] = useState<User[]>([]);
  const [searchQuery, setSearchQuery] = useState('');
  const [isLoading, setIsLoading] = useState(true);
  const [isExcelModalOpen, setIsExcelModalOpen] = useState(false);

  // Add / Edit Modal state
  const [isFormModalOpen, setIsFormModalOpen] = useState(false);
  const [editingTeacher, setEditingTeacher] = useState<User | null>(null);
  const [formData, setFormData] = useState({
    name: '',
    employeeId: '',
    password: '',
  });
  const [phoneList, setPhoneList] = useState<string[]>([]);
  const [phoneInput, setPhoneInput] = useState('');
  const [emailList, setEmailList] = useState<string[]>([]);
  const [emailInput, setEmailInput] = useState('');

  const [showEditPassword, setShowEditPassword] = useState(false);
  const [formError, setFormError] = useState<string | null>(null);
  const [isSaving, setIsSaving] = useState(false);

  // View Full Record Detail Modal (matching Student Master)
  const [detailTeacher, setDetailTeacher] = useState<User | null>(null);

  // Credentials View Dialog
  const [credentialsModalUser, setCredentialsModalUser] = useState<User | null>(null);
  const [copiedKey, setCopiedKey] = useState<string | null>(null);
  const [showPassword, setShowPassword] = useState(false);

  // Delete confirm state
  const [deletingId, setDeletingId] = useState<string | null>(null);
  const [deleteError, setDeleteError] = useState<string | null>(null);

  // Tag list managers for multiple phone numbers and email IDs
  const addContactTag = (
    rawInput: string,
    list: string[],
    setList: React.Dispatch<React.SetStateAction<string[]>>,
    setInput: React.Dispatch<React.SetStateAction<string>>
  ) => {
    if (!rawInput.trim()) return;
    const items = rawInput
      .split(',')
      .map((s) => s.trim())
      .filter((s) => s.length > 0 && !list.includes(s));
    if (items.length > 0) {
      setList([...list, ...items]);
    }
    setInput('');
  };

  const removeContactTag = (
    index: number,
    list: string[],
    setList: React.Dispatch<React.SetStateAction<string[]>>
  ) => {
    setList(list.filter((_, i) => i !== index));
  };

  const fetchTeachers = async () => {
    setIsLoading(true);
    try {
      const data = await api.getTeachers();
      setTeachers(data);
    } catch (err) {
      console.error('Failed to load teachers:', err);
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    fetchTeachers();

    // Listen for database re-sync events from Header or Auth
    const handleSync = () => fetchTeachers();
    window.addEventListener('database-resynced', handleSync);
    return () => window.removeEventListener('database-resynced', handleSync);
  }, []);

  const openAddModal = () => {
    setEditingTeacher(null);
    setFormData({
      name: '',
      employeeId: '',
      password: '',
    });
    setPhoneList([]);
    setPhoneInput('');
    setEmailList([]);
    setEmailInput('');
    setFormError(null);
    setIsFormModalOpen(true);
  };

  const openEditModal = (teacher: User) => {
    setEditingTeacher(teacher);
    setFormData({
      name: teacher.name || '',
      employeeId: teacher.employeeId || '',
      password: teacher.password || '',
    });
    setPhoneList(parseCommaList(teacher.contactNumber));
    setPhoneInput('');
    setEmailList(parseCommaList(teacher.emailId));
    setEmailInput('');
    setShowEditPassword(false);
    setFormError(null);
    setIsFormModalOpen(true);
  };

  const handleFormSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!formData.name.trim() || !formData.employeeId.trim()) {
      setFormError('Name and Employee ID are required.');
      return;
    }

    setIsSaving(true);
    setFormError(null);

    // Merge any pending unadded text input
    let finalPhones = [...phoneList];
    if (phoneInput.trim()) {
      const morePhones = phoneInput
        .split(',')
        .map((s) => s.trim())
        .filter((s) => s.length > 0 && !finalPhones.includes(s));
      finalPhones.push(...morePhones);
    }

    let finalEmails = [...emailList];
    if (emailInput.trim()) {
      const moreEmails = emailInput
        .split(',')
        .map((s) => s.trim())
        .filter((s) => s.length > 0 && !finalEmails.includes(s));
      finalEmails.push(...moreEmails);
    }

    const payload = {
      name: formData.name.trim(),
      employeeId: formData.employeeId.trim(),
      contactNumber: finalPhones.length > 0 ? finalPhones.join(', ') : undefined,
      emailId: finalEmails.length > 0 ? finalEmails.join(', ') : undefined,
    };

    try {
      if (editingTeacher) {
        // Admin alters password or details; username is immutable
        await api.updateTeacher(editingTeacher.id, {
          ...payload,
          password: formData.password.trim() || undefined,
        });
      } else {
        // System automatically generates unique username and password
        await api.createTeacher(payload);
      }
      setIsFormModalOpen(false);
      fetchTeachers();
    } catch (err: any) {
      setFormError(err.message || 'Operation failed');
    } finally {
      setIsSaving(false);
    }
  };

  const handleDelete = async (id: string) => {
    setDeleteError(null);
    try {
      await api.deleteTeacher(id);
      setDeletingId(null);
      fetchTeachers();
    } catch (err: any) {
      setDeleteError(err.message || 'Failed to delete teacher');
    }
  };

  const handleCopy = (text: string, type: string) => {
    navigator.clipboard.writeText(text);
    setCopiedKey(type);
    setTimeout(() => setCopiedKey(null), 2000);
  };

  const handleRegeneratePassword = () => {
    const newPwd = generateRandomPassword(9);
    setFormData((prev) => ({ ...prev, password: newPwd }));
    setShowEditPassword(true);
  };

  const previewUsername = () => {
    const cleanName = (formData.name || 'teacher')
      .toLowerCase()
      .trim()
      .replace(/[^a-z0-9]/g, '')
      .slice(0, 10);
    const cleanEmp = (formData.employeeId || 'empid')
      .toLowerCase()
      .trim()
      .replace(/[^a-z0-9]/g, '');
    return `${cleanName}_${cleanEmp}`;
  };

  const filteredTeachers = teachers.filter((t) => {
    const q = searchQuery.toLowerCase().trim();
    return (
      t.name.toLowerCase().includes(q) ||
      (t.employeeId && t.employeeId.toLowerCase().includes(q)) ||
      t.username.toLowerCase().includes(q) ||
      (t.emailId && t.emailId.toLowerCase().includes(q)) ||
      (t.contactNumber && t.contactNumber.toLowerCase().includes(q))
    );
  });

  return (
    <div className="space-y-6">
      {/* Top Banner / Actions Bar */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 bg-slate-900/60 p-5 rounded-2xl border border-slate-800">
        <div>
          <div className="flex items-center space-x-2.5">
            <div className="p-2 rounded-xl bg-purple-950/80 border border-purple-800/60 text-purple-300">
              <Users className="w-5 h-5" />
            </div>
            <div>
              <h2 className="text-lg font-bold text-slate-100">Teacher Profile Management</h2>
              <p className="text-xs text-slate-400">
                System-assigned unique usernames, randomized credentials, and password administration
              </p>
            </div>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-2.5">
          <button
            type="button"
            onClick={downloadTeacherTemplateExcel}
            title="Download teacher_schema_template.xlsx for Teacher Profile"
            className="inline-flex items-center space-x-1.5 px-3.5 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs font-medium border border-slate-700/80 transition-colors shadow-sm cursor-pointer"
          >
            <Download className="w-3.5 h-3.5 text-indigo-400" />
            <span>Download Schema Excel</span>
          </button>

          <button
            onClick={() => setIsExcelModalOpen(true)}
            className="inline-flex items-center space-x-1.5 px-3.5 py-2 rounded-xl bg-indigo-950/80 hover:bg-indigo-900/80 text-indigo-200 text-xs font-medium border border-indigo-800/70 transition-colors shadow-sm"
          >
            <FileSpreadsheet className="w-3.5 h-3.5 text-indigo-400" />
            <span>Upload Teacher Excel</span>
          </button>

          <button
            onClick={openAddModal}
            className="inline-flex items-center space-x-1.5 px-4 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-semibold shadow-md shadow-indigo-600/20 transition-all cursor-pointer"
          >
            <UserPlus className="w-3.5 h-3.5" />
            <span>Add Teacher Profile</span>
          </button>
        </div>
      </div>

      {/* Metric Cards */}
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
        <div className="bg-slate-900/60 border border-slate-800 rounded-xl p-4">
          <span className="text-xs text-slate-400 block">Total Faculty Profiles</span>
          <span className="text-2xl font-bold text-slate-100 mt-1 block">{teachers.length}</span>
        </div>
        <div className="bg-slate-900/60 border border-slate-800 rounded-xl p-4">
          <span className="text-xs text-slate-400 block">System-Assigned Usernames</span>
          <span className="text-2xl font-bold text-indigo-300 mt-1 block">100% Unique</span>
        </div>
        <div className="bg-slate-900/60 border border-slate-800 rounded-xl p-4">
          <span className="text-xs text-slate-400 block">Persistence Backend</span>
          <span className="text-2xl font-bold text-emerald-400 mt-1 block">Firestore Synced</span>
        </div>
      </div>

      {/* Search & Filter Bar */}
      <div className="bg-slate-900/40 p-4 rounded-xl border border-slate-800/80 flex items-center justify-between gap-4">
        <div className="relative flex-1 max-w-md">
          <Search className="w-4 h-4 text-slate-500 absolute left-3.5 top-1/2 -translate-y-1/2" />
          <input
            type="text"
            placeholder="Search by name, employee ID, username, or email..."
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            className="w-full pl-9 pr-4 py-2 bg-slate-950/60 border border-slate-800 rounded-xl text-xs text-slate-200 placeholder-slate-500 focus:outline-none focus:border-indigo-500"
          />
        </div>

        <button
          onClick={fetchTeachers}
          title="Refresh List"
          className="p-2 text-slate-400 hover:text-slate-200 hover:bg-slate-800 rounded-lg transition-colors cursor-pointer"
        >
          <RefreshCw className={`w-4 h-4 ${isLoading ? 'animate-spin' : ''}`} />
        </button>
      </div>

      {/* Teachers Table */}
      <div className="bg-slate-900/60 border border-slate-800 rounded-2xl overflow-hidden shadow-xl">
        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs text-slate-300">
            <thead className="bg-slate-950/70 border-b border-slate-800 text-slate-400 font-medium uppercase tracking-wider text-[11px]">
              <tr>
                <th className="px-5 py-3.5">Teacher Name</th>
                <th className="px-4 py-3.5">Employee ID</th>
                <th className="px-4 py-3.5">Assigned Login ID</th>
                <th className="px-4 py-3.5">Contact Details</th>
                <th className="px-4 py-3.5">Password Control</th>
                <th className="px-5 py-3.5 text-right">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-800/60">
              {filteredTeachers.length === 0 ? (
                <tr>
                  <td colSpan={6} className="px-6 py-12 text-center text-slate-500">
                    {searchQuery ? 'No teachers found matching your search.' : 'No teacher profiles created yet. Add one or upload via Excel.'}
                  </td>
                </tr>
              ) : (
                filteredTeachers.map((teacher) => (
                  <tr key={teacher.id} className="hover:bg-slate-800/40 transition-colors">
                    <td className="px-5 py-3.5 font-medium text-slate-100">
                      <div className="flex items-center space-x-2.5">
                        <div className="w-8 h-8 rounded-full bg-indigo-950 border border-indigo-800/60 text-indigo-300 flex items-center justify-center font-semibold text-xs">
                          {teacher.name.charAt(0).toUpperCase()}
                        </div>
                        <div>
                          <span className="block font-semibold text-slate-100">{teacher.name}</span>
                          <span className="text-[10px] text-slate-400">Role: Faculty</span>
                        </div>
                      </div>
                    </td>

                    <td className="px-4 py-3.5 font-mono text-indigo-300">
                      {teacher.employeeId || '—'}
                    </td>

                    <td className="px-4 py-3.5 font-mono text-slate-200">
                      <span className="px-2 py-0.5 rounded bg-slate-800 border border-slate-700/60 text-xs">
                        {teacher.username}
                      </span>
                    </td>

                    <td className="px-4 py-3.5">
                      {(() => {
                        const emails = parseCommaList(teacher.emailId);
                        const phones = parseCommaList(teacher.contactNumber);
                        if (emails.length === 0 && phones.length === 0) {
                          return <span className="text-slate-500 italic text-[11px]">No contact provided</span>;
                        }
                        return (
                          <div className="space-y-1.5">
                            {phones.length > 0 && (
                              <div className="flex flex-wrap gap-1 items-center">
                                <Phone className="w-3 h-3 text-slate-500 shrink-0" />
                                {phones.map((ph, idx) => (
                                  <span
                                    key={idx}
                                    className="px-1.5 py-0.5 rounded bg-slate-800 text-[10px] text-slate-300 font-mono border border-slate-700/60"
                                    title={ph}
                                  >
                                    {ph}
                                  </span>
                                ))}
                              </div>
                            )}
                            {emails.length > 0 && (
                              <div className="flex flex-wrap gap-1 items-center">
                                <Mail className="w-3 h-3 text-slate-500 shrink-0" />
                                {emails.map((em, idx) => (
                                  <span
                                    key={idx}
                                    className="px-1.5 py-0.5 rounded bg-slate-800 text-[10px] text-slate-300 font-mono truncate max-w-[170px] border border-slate-700/60"
                                    title={em}
                                  >
                                    {em}
                                  </span>
                                ))}
                              </div>
                            )}
                          </div>
                        );
                      })()}
                    </td>

                    <td className="px-4 py-3.5">
                      <div className="flex items-center space-x-2">
                        <button
                          onClick={() => {
                            setCredentialsModalUser(teacher);
                            setShowPassword(true);
                          }}
                          className="inline-flex items-center space-x-1.5 px-2.5 py-1 rounded-lg bg-indigo-950/60 hover:bg-indigo-900/60 text-indigo-300 border border-indigo-800/50 text-[11px] font-mono transition-colors cursor-pointer"
                        >
                          <Key className="w-3 h-3" />
                          <span>View Password</span>
                        </button>
                        <button
                          onClick={() => openEditModal(teacher)}
                          className="inline-flex items-center space-x-1 px-2.5 py-1 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-300 text-[11px] transition-colors cursor-pointer"
                        >
                          <Edit2 className="w-3 h-3 text-slate-400" />
                          <span>Alter</span>
                        </button>
                      </div>
                    </td>

                    <td className="px-5 py-3.5 text-right">
                      <div className="flex items-center justify-end space-x-1.5">
                        <button
                          onClick={() => setDetailTeacher(teacher)}
                          title="View Full Record & Contact Channels"
                          className="p-1.5 text-slate-400 hover:text-indigo-300 hover:bg-slate-800 rounded-lg transition-colors cursor-pointer"
                        >
                          <Eye className="w-3.5 h-3.5" />
                        </button>
                        <button
                          onClick={() => openEditModal(teacher)}
                          title="Edit Profile & Password"
                          className="p-1.5 text-slate-400 hover:text-indigo-300 hover:bg-slate-800 rounded-lg transition-colors cursor-pointer"
                        >
                          <Edit2 className="w-3.5 h-3.5" />
                        </button>
                        <button
                          onClick={() => setDeletingId(teacher.id)}
                          title="Delete Profile"
                          className="p-1.5 text-slate-400 hover:text-rose-400 hover:bg-rose-950/40 rounded-lg transition-colors cursor-pointer"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Add / Edit Profile Modal */}
      {isFormModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-sm">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-md shadow-2xl p-6">
            <h3 className="text-base font-bold text-slate-100 mb-1">
              {editingTeacher ? 'Alter Teacher Profile & Password' : 'Create New Teacher Profile'}
            </h3>
            <p className="text-xs text-slate-400 mb-4">
              {editingTeacher
                ? 'Username is immutable. Admin has option to view or alter password.'
                : 'Login ID and password will automatically be assigned by the system.'}
            </p>

            {formError && (
              <div className="mb-4 p-3 rounded-xl bg-rose-950/50 border border-rose-800/60 text-rose-300 text-xs flex items-center space-x-2">
                <AlertCircle className="w-4 h-4 shrink-0" />
                <span>{formError}</span>
              </div>
            )}

            <form onSubmit={handleFormSubmit} className="space-y-3.5">
              <div>
                <label className="block text-xs font-medium text-slate-300 mb-1">
                  Full Name *
                </label>
                <input
                  type="text"
                  required
                  placeholder="e.g. Dr. Robert Miller"
                  value={formData.name}
                  onChange={(e) => setFormData({ ...formData, name: e.target.value })}
                  className="w-full px-3 py-2 bg-slate-950 border border-slate-800 rounded-xl text-xs text-slate-200 focus:outline-none focus:border-indigo-500"
                />
              </div>

              <div>
                <label className="block text-xs font-medium text-slate-300 mb-1">
                  Employee ID *
                </label>
                <input
                  type="text"
                  required
                  placeholder="e.g. EMP-1042"
                  value={formData.employeeId}
                  onChange={(e) => setFormData({ ...formData, employeeId: e.target.value })}
                  className="w-full px-3 py-2 bg-slate-950 border border-slate-800 rounded-xl text-xs text-slate-200 focus:outline-none focus:border-indigo-500 font-mono"
                />
              </div>

              {/* Automatic System Assignment Section */}
              {!editingTeacher ? (
                <div className="p-3.5 rounded-xl bg-indigo-950/40 border border-indigo-800/50 space-y-2 text-xs">
                  <div className="flex items-center space-x-1.5 text-indigo-300 font-semibold">
                    <Sparkles className="w-3.5 h-3.5 text-indigo-400" />
                    <span>System Automated Credentials</span>
                  </div>
                  <div className="space-y-1 text-slate-400 text-[11px]">
                    <div className="flex justify-between items-center py-0.5">
                      <span>Unique Login ID:</span>
                      <span className="font-mono text-indigo-300 font-semibold bg-slate-950 px-2 py-0.5 rounded border border-slate-800">
                        {previewUsername()} (Auto-Guaranteed)
                      </span>
                    </div>
                    <div className="flex justify-between items-center py-0.5">
                      <span>Password:</span>
                      <span className="font-mono text-emerald-400 font-semibold bg-slate-950 px-2 py-0.5 rounded border border-slate-800">
                        Auto-Generated Randomly
                      </span>
                    </div>
                  </div>
                </div>
              ) : (
                /* Edit Mode: Admin Alters/Views Password, Username is Immutable */
                <div className="space-y-3 pt-1">
                  <div>
                    <label className="block text-xs font-medium text-slate-400 mb-1">
                      Assigned Login ID (Immutable)
                    </label>
                    <input
                      type="text"
                      disabled
                      value={editingTeacher.username}
                      className="w-full px-3 py-2 bg-slate-950/80 border border-slate-800/60 rounded-xl text-xs text-slate-400 font-mono cursor-not-allowed"
                    />
                  </div>

                  <div>
                    <div className="flex items-center justify-between mb-1">
                      <label className="text-xs font-medium text-slate-200">
                        Password (Alter / View)
                      </label>
                      <button
                        type="button"
                        onClick={handleRegeneratePassword}
                        className="text-[11px] text-indigo-400 hover:text-indigo-300 flex items-center space-x-1"
                      >
                        <RefreshCw className="w-3 h-3" />
                        <span>Generate Random</span>
                      </button>
                    </div>

                    <div className="relative">
                      <input
                        type={showEditPassword ? 'text' : 'password'}
                        required
                        placeholder="Alter account password..."
                        value={formData.password}
                        onChange={(e) => setFormData({ ...formData, password: e.target.value })}
                        className="w-full pl-3 pr-10 py-2 bg-slate-950 border border-slate-800 rounded-xl text-xs text-slate-200 focus:outline-none focus:border-indigo-500 font-mono"
                      />
                      <button
                        type="button"
                        onClick={() => setShowEditPassword(!showEditPassword)}
                        className="absolute inset-y-0 right-0 pr-3 flex items-center text-slate-400 hover:text-slate-200"
                      >
                        {showEditPassword ? <EyeOff className="w-3.5 h-3.5" /> : <Eye className="w-3.5 h-3.5" />}
                      </button>
                    </div>
                  </div>
                </div>
              )}

              {/* Faculty Multi-Contact Section (Identical to Student Master) */}
              <div className="p-4 rounded-xl bg-slate-950/60 border border-slate-800 space-y-3.5">
                <div className="flex items-center justify-between">
                  <div className="text-[11px] font-semibold text-slate-400 uppercase tracking-wider">
                    Faculty Contact Channels (Multiple Allowed)
                  </div>
                  <span className="text-[10px] text-slate-500">Comma-separated or press Add</span>
                </div>

                {/* Multiple Phone Numbers */}
                <div>
                  <label className="block text-xs font-medium text-slate-300 mb-1">
                    Phone Number(s)
                  </label>
                  <div className="flex gap-2">
                    <input
                      type="text"
                      placeholder="e.g. +91-90160-86894, +91-72824-87704"
                      value={phoneInput}
                      onChange={(e) => setPhoneInput(e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') {
                          e.preventDefault();
                          addContactTag(phoneInput, phoneList, setPhoneList, setPhoneInput);
                        }
                      }}
                      className="flex-1 px-3 py-2 bg-slate-900 border border-slate-800 rounded-xl text-xs text-slate-200 focus:outline-none focus:border-indigo-500 font-mono"
                    />
                    <button
                      type="button"
                      onClick={() => addContactTag(phoneInput, phoneList, setPhoneList, setPhoneInput)}
                      className="px-3.5 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-xs font-medium text-slate-200 cursor-pointer"
                    >
                      Add
                    </button>
                  </div>
                  {phoneList.length > 0 && (
                    <div className="flex flex-wrap gap-1.5 mt-2">
                      {phoneList.map((ph, idx) => (
                        <span
                          key={idx}
                          className="inline-flex items-center space-x-1 px-2.5 py-1 rounded-lg bg-indigo-950/70 border border-indigo-800/60 text-indigo-300 text-xs font-mono"
                        >
                          <span>{ph}</span>
                          <button
                            type="button"
                            onClick={() => removeContactTag(idx, phoneList, setPhoneList)}
                            className="text-slate-400 hover:text-rose-400 cursor-pointer"
                          >
                            <X className="w-3 h-3" />
                          </button>
                        </span>
                      ))}
                    </div>
                  )}
                </div>

                {/* Multiple Emails */}
                <div>
                  <label className="block text-xs font-medium text-slate-300 mb-1">
                    Email Address(es)
                  </label>
                  <div className="flex gap-2">
                    <input
                      type="text"
                      placeholder="e.g. aditi.iyer@example.edu, aditi.iyer1@example.edu"
                      value={emailInput}
                      onChange={(e) => setEmailInput(e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') {
                          e.preventDefault();
                          addContactTag(emailInput, emailList, setEmailList, setEmailInput);
                        }
                      }}
                      className="flex-1 px-3 py-2 bg-slate-900 border border-slate-800 rounded-xl text-xs text-slate-200 focus:outline-none focus:border-indigo-500"
                    />
                    <button
                      type="button"
                      onClick={() => addContactTag(emailInput, emailList, setEmailList, setEmailInput)}
                      className="px-3.5 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-xs font-medium text-slate-200 cursor-pointer"
                    >
                      Add
                    </button>
                  </div>
                  {emailList.length > 0 && (
                    <div className="flex flex-wrap gap-1.5 mt-2">
                      {emailList.map((em, idx) => (
                        <span
                          key={idx}
                          className="inline-flex items-center space-x-1 px-2.5 py-1 rounded-lg bg-slate-800 text-slate-200 text-xs border border-slate-700/60"
                        >
                          <span className="truncate max-w-[200px]" title={em}>{em}</span>
                          <button
                            type="button"
                            onClick={() => removeContactTag(idx, emailList, setEmailList)}
                            className="text-slate-400 hover:text-rose-400 cursor-pointer"
                          >
                            <X className="w-3 h-3" />
                          </button>
                        </span>
                      ))}
                    </div>
                  )}
                </div>
              </div>

              <div className="flex justify-end space-x-2.5 pt-3">
                <button
                  type="button"
                  onClick={() => setIsFormModalOpen(false)}
                  className="px-4 py-2 rounded-xl text-xs font-medium text-slate-400 hover:bg-slate-800 cursor-pointer"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  disabled={isSaving}
                  className="px-4 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-semibold shadow-md shadow-indigo-600/20 disabled:opacity-50 cursor-pointer"
                >
                  {isSaving ? 'Saving...' : editingTeacher ? 'Save Alterations' : 'Assign & Create Profile'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* View Full Faculty Record Detail Modal (Identical to Student Master) */}
      {detailTeacher && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-sm">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-lg shadow-2xl p-6 space-y-4">
            <div className="flex items-center justify-between border-b border-slate-800 pb-3">
              <div className="flex items-center space-x-3">
                <div className="w-10 h-10 rounded-xl bg-indigo-950 text-indigo-300 border border-indigo-800/60 flex items-center justify-center font-bold text-base">
                  {detailTeacher.name.charAt(0).toUpperCase()}
                </div>
                <div>
                  <h3 className="font-bold text-slate-100 text-sm">{detailTeacher.name}</h3>
                  <div className="flex items-center space-x-2 mt-0.5">
                    <span className="text-xs text-indigo-400 font-mono">
                      EMP ID: {detailTeacher.employeeId}
                    </span>
                    <span className="text-slate-600">•</span>
                    <span className="text-xs text-slate-400 font-mono">
                      User: {detailTeacher.username}
                    </span>
                  </div>
                </div>
              </div>
              <button
                onClick={() => setDetailTeacher(null)}
                className="text-slate-400 hover:text-slate-200 cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-3.5 text-xs">
              <div className="grid grid-cols-2 gap-3 bg-slate-950/60 p-3 rounded-xl border border-slate-800">
                <div>
                  <span className="text-slate-400 block text-[11px]">System Role</span>
                  <span className="text-emerald-400 font-semibold uppercase">Faculty (Teacher)</span>
                </div>
                <div>
                  <span className="text-slate-400 block text-[11px]">Assigned Login ID</span>
                  <span className="text-slate-200 font-mono font-semibold">{detailTeacher.username}</span>
                </div>
                <div className="col-span-2 pt-2 border-t border-slate-800/60 flex items-center justify-between">
                  <div>
                    <span className="text-slate-400 block text-[11px]">Current Password</span>
                    <span className="text-emerald-400 font-mono font-semibold text-sm">
                      {detailTeacher.password || 'Not set'}
                    </span>
                  </div>
                  {detailTeacher.password && (
                    <button
                      onClick={() => handleCopy(detailTeacher.password!, 'detail_pwd')}
                      className="inline-flex items-center space-x-1 px-2.5 py-1 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-300 text-[11px] font-mono cursor-pointer"
                      title="Copy current password"
                    >
                      {copiedKey === 'detail_pwd' ? (
                        <Check className="w-3.5 h-3.5 text-emerald-400" />
                      ) : (
                        <Copy className="w-3.5 h-3.5 text-slate-400" />
                      )}
                      <span>Copy</span>
                    </button>
                  )}
                </div>
              </div>

              {/* Phone numbers */}
              <div>
                <span className="text-slate-400 block mb-1.5 text-[11px] font-semibold flex items-center space-x-1.5">
                  <Phone className="w-3 h-3 text-slate-500" />
                  <span>Registered Contact Numbers</span>
                </span>
                {(() => {
                  const phones = parseCommaList(detailTeacher.contactNumber);
                  return phones.length > 0 ? (
                    <div className="flex flex-wrap gap-1.5">
                      {phones.map((p, i) => (
                        <span key={i} className="px-2.5 py-1 rounded-lg bg-indigo-950/70 text-indigo-300 font-mono text-xs border border-indigo-800/60">
                          {p}
                        </span>
                      ))}
                    </div>
                  ) : (
                    <span className="text-slate-500 italic text-[11px]">None provided</span>
                  );
                })()}
              </div>

              {/* Email IDs */}
              <div>
                <span className="text-slate-400 block mb-1.5 text-[11px] font-semibold flex items-center space-x-1.5">
                  <Mail className="w-3 h-3 text-slate-500" />
                  <span>Registered Email Addresses</span>
                </span>
                {(() => {
                  const emails = parseCommaList(detailTeacher.emailId);
                  return emails.length > 0 ? (
                    <div className="flex flex-wrap gap-1.5">
                      {emails.map((e, i) => (
                        <span key={i} className="px-2.5 py-1 rounded-lg bg-slate-800 text-slate-200 text-xs border border-slate-700/60 font-mono">
                          {e}
                        </span>
                      ))}
                    </div>
                  ) : (
                    <span className="text-slate-500 italic text-[11px]">None provided</span>
                  );
                })()}
              </div>
            </div>

            <div className="flex justify-between items-center pt-3 border-t border-slate-800">
              <button
                onClick={() => {
                  const target = detailTeacher;
                  setDetailTeacher(null);
                  setCredentialsModalUser(target);
                  setShowPassword(true);
                }}
                className="inline-flex items-center space-x-1.5 text-xs text-indigo-400 hover:text-indigo-300 font-medium cursor-pointer"
              >
                <Key className="w-3.5 h-3.5" />
                <span>View Login Credentials</span>
              </button>
              <button
                onClick={() => setDetailTeacher(null)}
                className="px-4 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-xs font-medium text-slate-200 cursor-pointer"
              >
                Close
              </button>
            </div>
          </div>
        </div>
      )}

      {/* View Credentials Modal */}
      {credentialsModalUser && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-sm">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-sm shadow-2xl p-6">
            <div className="flex items-center space-x-2.5 mb-3">
              <div className="p-2 rounded-xl bg-indigo-950 border border-indigo-800/60 text-indigo-400">
                <Key className="w-4 h-4" />
              </div>
              <h3 className="font-bold text-slate-100 text-sm">Faculty Access Credentials</h3>
            </div>
            <p className="text-xs text-slate-400 mb-4">
              Credentials for <strong className="text-slate-200">{credentialsModalUser.name}</strong> ({credentialsModalUser.employeeId})
            </p>

            <div className="space-y-3 bg-slate-950/70 p-4 rounded-xl border border-slate-800">
              <div>
                <label className="text-[11px] text-slate-400 block mb-1">Assigned Username</label>
                <div className="flex items-center justify-between bg-slate-900 px-3 py-2 rounded-lg border border-slate-800 font-mono text-xs text-slate-200">
                  <span>{credentialsModalUser.username}</span>
                  <button
                    onClick={() => handleCopy(credentialsModalUser.username, 'username')}
                    className="text-slate-400 hover:text-indigo-400 cursor-pointer"
                    title="Copy username"
                  >
                    {copiedKey === 'username' ? <Check className="w-3.5 h-3.5 text-emerald-400" /> : <Copy className="w-3.5 h-3.5" />}
                  </button>
                </div>
              </div>

              <div>
                <label className="text-[11px] text-slate-400 block mb-1">Current Password</label>
                <div className="flex items-center justify-between bg-slate-900 px-3 py-2 rounded-lg border border-slate-800 font-mono text-xs text-slate-200">
                  <span className="font-semibold text-emerald-400 select-all tracking-wide">
                    {showPassword ? (credentialsModalUser.password || 'Not set') : '••••••••••••'}
                  </span>
                  <div className="flex items-center space-x-2">
                    <button
                      onClick={() => setShowPassword(!showPassword)}
                      className="text-slate-400 hover:text-slate-200 cursor-pointer p-1"
                      title={showPassword ? 'Hide password' : 'Show password'}
                    >
                      {showPassword ? <EyeOff className="w-3.5 h-3.5" /> : <Eye className="w-3.5 h-3.5" />}
                    </button>
                    {credentialsModalUser.password && (
                      <button
                        onClick={() => handleCopy(credentialsModalUser.password!, 'password')}
                        className="text-slate-400 hover:text-indigo-400 cursor-pointer p-1"
                        title="Copy password"
                      >
                        {copiedKey === 'password' ? <Check className="w-3.5 h-3.5 text-emerald-400" /> : <Copy className="w-3.5 h-3.5" />}
                      </button>
                    )}
                  </div>
                </div>
              </div>
            </div>

            <div className="flex justify-between items-center mt-5">
              <button
                onClick={() => {
                  const target = credentialsModalUser;
                  setCredentialsModalUser(null);
                  openEditModal(target);
                }}
                className="text-xs text-indigo-400 hover:text-indigo-300 font-medium"
              >
                Alter Password
              </button>
              <button
                onClick={() => setCredentialsModalUser(null)}
                className="px-4 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-xs font-medium text-slate-200 cursor-pointer"
              >
                Close
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Delete Confirmation Modal */}
      {deletingId && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-sm">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-sm shadow-2xl p-6">
            <h3 className="font-bold text-slate-100 text-sm mb-2">Confirm Profile Removal</h3>
            <p className="text-xs text-slate-400 mb-4">
              Are you sure you want to remove this teacher profile? Their login access will be revoked immediately from the persistent database.
            </p>
            {deleteError && (
              <div className="mb-4 p-3 rounded-xl bg-rose-950/50 border border-rose-800/60 text-rose-300 text-xs flex items-center space-x-2">
                <AlertCircle className="w-4 h-4 shrink-0 text-rose-400" />
                <span>{deleteError}</span>
              </div>
            )}
            <div className="flex justify-end space-x-2.5">
              <button
                onClick={() => {
                  setDeletingId(null);
                  setDeleteError(null);
                }}
                className="px-4 py-2 rounded-xl text-xs font-medium text-slate-400 hover:bg-slate-800 cursor-pointer"
              >
                Cancel
              </button>
              <button
                onClick={() => handleDelete(deletingId)}
                className="px-4 py-2 rounded-xl bg-rose-600 hover:bg-rose-500 text-white text-xs font-semibold shadow-md shadow-rose-600/20 cursor-pointer"
              >
                Delete Profile
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Bulk Excel Upload Modal with Error Debugging */}
      <ExcelUploadModal<User>
        isOpen={isExcelModalOpen}
        onClose={() => setIsExcelModalOpen(false)}
        title="Upload Teacher Excel Data"
        description="Extract teacher details, automatically generate unique user ID and random passwords, and persist to Firebase"
        schemaDownloadUrl={api.downloadTeacherSchemaUrl()}
        schemaDownloadLabel="Download Teacher Schema"
        onDownloadSchema={downloadTeacherTemplateExcel}
        onUpload={api.uploadTeacherExcel}
        onSuccess={fetchTeachers}
      />
    </div>
  );
};
