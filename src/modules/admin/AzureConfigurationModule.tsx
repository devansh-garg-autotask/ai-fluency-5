import React, { useState, useEffect, useCallback } from 'react';
import {
  Settings,
  Key,
  Copy,
  Check,
  CheckCircle2,
  AlertCircle,
  AlertTriangle,
  RefreshCw,
  Database,
  ExternalLink,
  ShieldCheck,
  Eye,
  EyeOff,
  Cloud,
} from 'lucide-react';
import {
  getOutlookConfig,
  saveOutlookCredentials,
} from '../../services/outlookApi.ts';

export const AzureConfigurationModule: React.FC = () => {
  const [isConfigured, setIsConfigured] = useState<boolean>(false);
  const [redirectUri, setRedirectUri] = useState<string>('');
  const [clientIdMasked, setClientIdMasked] = useState<string>('');
  const [isLoading, setIsLoading] = useState<boolean>(true);

  // Form state
  const [cfgClientId, setCfgClientId] = useState<string>('');
  const [cfgClientSecret, setCfgClientSecret] = useState<string>('');
  const [showSecret, setShowSecret] = useState<boolean>(false);
  const [isSaving, setIsSaving] = useState<boolean>(false);
  const [copiedRedirect, setCopiedRedirect] = useState<boolean>(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [successMessage, setSuccessMessage] = useState<string | null>(null);

  const loadConfig = useCallback(async () => {
    setIsLoading(true);
    setErrorMessage(null);
    try {
      const cfg = await getOutlookConfig();
      setIsConfigured(cfg.isConfigured);
      setRedirectUri(cfg.redirectUri);
      if (cfg.clientIdMasked) {
        setClientIdMasked(cfg.clientIdMasked);
      }
    } catch (err: any) {
      console.error('Failed to load Azure configuration:', err);
      setErrorMessage(err.message || 'Unable to retrieve current Azure configuration');
    } finally {
      setIsLoading(false);
    }
  }, []);

  useEffect(() => {
    loadConfig();
  }, [loadConfig]);

  const copyRedirectUri = () => {
    const uriToCopy = redirectUri || `${window.location.origin}/api/outlook/callback`;
    navigator.clipboard.writeText(uriToCopy);
    setCopiedRedirect(true);
    setTimeout(() => setCopiedRedirect(false), 2500);
  };

  const handleSave = async (e: React.FormEvent) => {
    e.preventDefault();
    setErrorMessage(null);
    setSuccessMessage(null);

    const cleanId = cfgClientId.trim();
    if (!cleanId) {
      setErrorMessage('Application (Client) ID is required.');
      return;
    }

    setIsSaving(true);
    try {
      const res = await saveOutlookCredentials(cleanId, cfgClientSecret.trim());
      setSuccessMessage(res.message || 'Azure credentials saved and persisted to database!');
      setCfgClientSecret('');
      await loadConfig();
    } catch (err: any) {
      setErrorMessage(err.message || 'Failed to save Azure credentials');
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <div className="space-y-6">
      {/* Top Banner */}
      <div className="bg-slate-900/60 p-5 rounded-2xl border border-slate-800 flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="flex items-center space-x-3.5">
          <div className="w-12 h-12 rounded-xl bg-indigo-950/80 border border-indigo-800/60 text-indigo-300 flex items-center justify-center shrink-0 shadow-lg shadow-indigo-950/50">
            <Cloud className="w-6 h-6" />
          </div>
          <div>
            <div className="flex items-center space-x-2.5">
              <h2 className="text-lg font-bold text-slate-100">Microsoft Azure Configuration</h2>
              <span
                className={`text-[11px] px-2.5 py-0.5 rounded-full font-semibold border ${
                  isConfigured
                    ? 'bg-emerald-950/80 text-emerald-300 border-emerald-800/60'
                    : 'bg-amber-950/80 text-amber-300 border-amber-800/60'
                }`}
              >
                {isConfigured ? 'Active & Configured' : 'Credentials Needed'}
              </span>
            </div>
            <p className="text-xs text-slate-400 mt-0.5">
              Configure institutional Microsoft Entra ID (Azure) App Registration credentials for authentic Outlook OAuth integrations.
            </p>
          </div>
        </div>

        <div className="flex items-center space-x-3">
          <div className="flex items-center space-x-1.5 text-emerald-400 font-mono text-[11px] bg-emerald-950/40 px-3 py-1.5 rounded-xl border border-emerald-800/40">
            <Database className="w-3.5 h-3.5" />
            <span>Persistent Firestore DB</span>
          </div>

          <button
            onClick={loadConfig}
            disabled={isLoading}
            className="flex items-center space-x-1.5 px-3 py-1.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 text-xs font-medium transition-all cursor-pointer disabled:opacity-50"
            title="Refresh current configuration"
          >
            <RefreshCw className={`w-3.5 h-3.5 text-indigo-400 ${isLoading ? 'animate-spin' : ''}`} />
            <span>Refresh</span>
          </button>
        </div>
      </div>

      {/* Notifications */}
      {successMessage && (
        <div className="p-4 rounded-xl bg-emerald-950/50 border border-emerald-800/70 flex items-center space-x-3 text-xs text-emerald-300 shadow-lg shadow-emerald-950/20">
          <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0" />
          <span className="font-medium">{successMessage}</span>
        </div>
      )}

      {errorMessage && (
        <div className="p-4 rounded-xl bg-rose-950/50 border border-rose-800/70 flex items-center space-x-3 text-xs text-rose-300 shadow-lg shadow-rose-950/20">
          <AlertCircle className="w-4 h-4 text-rose-400 shrink-0" />
          <span className="font-medium">{errorMessage}</span>
        </div>
      )}

      {/* Main Grid: Settings & Guidance */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Left 2 Cols: Credentials Form */}
        <div className="lg:col-span-2 bg-slate-900/60 border border-slate-800 rounded-2xl p-6 space-y-6">
          <div className="flex items-center space-x-2.5 pb-4 border-b border-slate-800">
            <Key className="w-4 h-4 text-indigo-400" />
            <h3 className="text-sm font-bold text-slate-100">Application Credentials</h3>
          </div>

          {/* Current Status Preview */}
          <div className="p-4 rounded-xl bg-slate-950/70 border border-slate-800 text-xs space-y-3">
            <div className="flex items-center justify-between">
              <span className="text-slate-400 font-medium">Current Registered Client ID</span>
              <span className="font-mono text-indigo-300 font-semibold">
                {clientIdMasked || (isConfigured ? 'Configured in Environment' : 'Not set')}
              </span>
            </div>
            <div className="flex items-center justify-between border-t border-slate-800/60 pt-2.5 text-[11px] text-slate-400">
              <span>Client Secret Storage</span>
              <span className="text-slate-300">
                Optional (Public Client mode supported with PKCE)
              </span>
            </div>
          </div>

          {/* Form */}
          <form onSubmit={handleSave} className="space-y-4 text-xs">
            <div className="space-y-1.5">
              <div className="flex items-center justify-between">
                <label className="text-slate-300 font-semibold block">
                  Application (Client) ID *
                </label>
                <span className="text-[10px] text-emerald-400 font-mono bg-emerald-950/60 px-2 py-0.5 rounded border border-emerald-900/60">
                  Required
                </span>
              </div>
              <input
                type="text"
                value={cfgClientId}
                onChange={(e) => setCfgClientId(e.target.value)}
                placeholder="e.g. 00000000-0000-0000-0000-000000000000"
                required
                className="w-full px-3.5 py-2.5 rounded-xl bg-slate-950 border border-slate-800 text-slate-100 placeholder-slate-500 font-mono focus:outline-none focus:border-indigo-500 transition-colors"
              />
              <p className="text-[11px] text-slate-400">
                Found in Azure Portal &gt; App registrations &gt; Your App &gt; Overview.
              </p>
            </div>

            <div className="space-y-1.5">
              <div className="flex items-center justify-between">
                <label className="text-slate-300 font-semibold block">
                  Client Secret Value
                </label>
                <span className="text-[10px] text-slate-400 font-mono bg-slate-800/60 px-2 py-0.5 rounded border border-slate-700/60">
                  Optional (Public Client)
                </span>
              </div>
              <div className="relative">
                <input
                  type={showSecret ? 'text' : 'password'}
                  value={cfgClientSecret}
                  onChange={(e) => setCfgClientSecret(e.target.value)}
                  placeholder="Optional — leave blank for Public Client mode"
                  className="w-full pl-3.5 pr-10 py-2.5 rounded-xl bg-slate-950 border border-slate-800 text-slate-100 placeholder-slate-500 font-mono focus:outline-none focus:border-indigo-500 transition-colors"
                />
                <button
                  type="button"
                  onClick={() => setShowSecret(!showSecret)}
                  className="absolute inset-y-0 right-0 pr-3.5 flex items-center text-slate-500 hover:text-slate-300"
                >
                  {showSecret ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                </button>
              </div>
              <p className="text-[11px] text-slate-400 leading-relaxed">
                Leave empty when using Public Client mode. If using a Web Client with Client Secret, paste the Secret Value here.
              </p>
            </div>

            <div className="pt-3 border-t border-slate-800 flex items-center justify-between">
              <span className="text-[11px] text-slate-500 flex items-center space-x-1.5">
                <ShieldCheck className="w-3.5 h-3.5 text-indigo-400" />
                <span>Persisted to system_outlook_config in Firestore</span>
              </span>

              <button
                type="submit"
                disabled={isSaving}
                className="px-5 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-white font-semibold flex items-center space-x-2 cursor-pointer shadow-lg shadow-indigo-600/30 transition-all disabled:opacity-50 text-xs"
              >
                {isSaving ? (
                  <>
                    <RefreshCw className="w-4 h-4 animate-spin" />
                    <span>Saving to Database...</span>
                  </>
                ) : (
                  <>
                    <CheckCircle2 className="w-4 h-4" />
                    <span>Save &amp; Activate Configuration</span>
                  </>
                )}
              </button>
            </div>
          </form>
        </div>

        {/* Right Col: Setup Guidance & Redirect URI */}
        <div className="space-y-6">
          {/* Redirect URI Box */}
          <div className="bg-slate-900/60 border border-slate-800 rounded-2xl p-5 space-y-3">
            <span className="text-slate-300 font-bold block text-xs">
              OAuth 2.0 Web Redirect URI
            </span>
            <p className="text-[11px] text-slate-400">
              Add this callback URL to your Microsoft Entra ID App Registration under <strong className="text-slate-200">Authentication</strong>:
            </p>

            <div className="flex items-center justify-between bg-slate-950 p-2.5 rounded-xl border border-slate-800 font-mono text-[11px] text-indigo-300">
              <span className="truncate select-all pr-2">
                {redirectUri || `${window.location.origin}/api/outlook/callback`}
              </span>
              <button
                type="button"
                onClick={copyRedirectUri}
                className="p-1.5 text-slate-400 hover:text-white hover:bg-slate-800 rounded-lg transition-colors shrink-0 cursor-pointer"
                title="Copy Redirect URI"
              >
                {copiedRedirect ? <Check className="w-4 h-4 text-emerald-400" /> : <Copy className="w-4 h-4" />}
              </button>
            </div>
            {copiedRedirect && (
              <span className="text-[10px] text-emerald-400 font-medium block">
                Copied Redirect URI to clipboard!
              </span>
            )}
          </div>

          {/* Public Client Flow Recommendation */}
          <div className="bg-slate-900/60 border border-slate-800 rounded-2xl p-5 space-y-3">
            <div className="flex items-center space-x-2 text-emerald-400">
              <CheckCircle2 className="w-4 h-4 shrink-0" />
              <h4 className="font-bold text-xs">Public Client Mode (Recommended)</h4>
            </div>
            <p className="text-[11px] text-slate-300 leading-relaxed">
              In Azure Portal &gt; <strong className="text-white">Authentication</strong> &gt; <strong className="text-white">Advanced settings</strong>, set:
            </p>
            <div className="p-3 bg-slate-950 rounded-xl border border-slate-800 text-[11px] font-mono text-emerald-300">
              Allow public client flows: Yes
            </div>
            <p className="text-[11px] text-slate-400 leading-relaxed">
              This enables PKCE token exchange and device code flow (microsoft.com/devicelogin) without exposing client secrets in faculty browser sessions.
            </p>
          </div>

          {/* Scope Checklist */}
          <div className="bg-slate-900/60 border border-slate-800 rounded-2xl p-5 space-y-2.5 text-xs">
            <h4 className="font-bold text-slate-200 text-xs">Required API Permissions</h4>
            <p className="text-[11px] text-slate-400">
              Under Microsoft Graph &gt; Delegated permissions:
            </p>
            <ul className="space-y-1.5 text-[11px] font-mono text-indigo-300">
              <li className="flex items-center space-x-1.5">
                <span className="w-1.5 h-1.5 rounded-full bg-indigo-500" />
                <span>openid, profile, email</span>
              </li>
              <li className="flex items-center space-x-1.5">
                <span className="w-1.5 h-1.5 rounded-full bg-indigo-500" />
                <span>offline_access</span>
              </li>
              <li className="flex items-center space-x-1.5">
                <span className="w-1.5 h-1.5 rounded-full bg-indigo-500" />
                <span>User.Read</span>
              </li>
              <li className="flex items-center space-x-1.5">
                <span className="w-1.5 h-1.5 rounded-full bg-indigo-500" />
                <span>Mail.ReadWrite, Mail.Send</span>
              </li>
            </ul>
          </div>
        </div>
      </div>
    </div>
  );
};
