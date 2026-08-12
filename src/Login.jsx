import React, { useState } from 'react';
import { auth, db } from './firebase';
import logoSidamon from './assets/logo-sidamon.png';

const Icon = ({ name, size = 20, className = "" }) => {
    const paths = {
        "boxes": '<path d="M2.97 12.92A2 2 0 0 0 2 14.63v3.24a2 2 0 0 0 .97 1.71l3 1.8a2 2 0 0 0 2.06 0L12 19v-5.5l-5-3-4.03 2.42Z"/><path d="m7 16.5-4.74-2.85"/><path d="m7 16.5 5-3"/><path d="M7 16.5v5.17"/><path d="M12 13.5V19l3.97 2.38a2 2 0 0 0 2.06 0l3-1.8a2 2 0 0 0 .97-1.71v-3.24a2 2 0 0 0-.97-1.71L17 10.5l-5 3Z"/><path d="m17 16.5-5-3"/><path d="m17 16.5 4.74-2.85"/><path d="M17 16.5v5.17"/><path d="M7.97 4.42A2 2 0 0 0 7 6.13v4.37l5 3 5-3V6.13a2 2 0 0 0-.97-1.71l-3-1.8a2 2 0 0 0-2.06 0l-3 1.8Z"/><path d="m12 8 4.74-2.85"/><path d="m12 8-4.74-2.85"/><path d="M12 8v5.5"/>',
        "user": '<path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/>',
        "mail": '<rect width="20" height="16" x="2" y="4" rx="2"/><path d="m22 7-8.97 5.7a1.94 1.94 0 0 1-2.06 0L2 7"/>',
        "lock": '<rect width="18" height="11" x="3" y="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/>',
        "loader-2": '<path d="M21 12a9 9 0 1 1-6.219-8.56"/>',
        "alert-circle": '<circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>',
        "check-circle-2": '<path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/>'
    };
    return (
        <svg xmlns="http://www.w3.org/2000/svg" width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className} dangerouslySetInnerHTML={{ __html: paths[name] || '' }} />
    );
};


export default function Login() {
    const [isLogin, setIsLogin] = useState(true);
    const [email, setEmail] = useState('');
    const [password, setPassword] = useState('');
    const [name, setName] = useState(''); // Hanya untuk Sign Up
    const [error, setError] = useState('');
    const [loading, setLoading] = useState(false);
    const [successMsg, setSuccessMsg] = useState('');

    const handleSubmit = async (e) => {
        e.preventDefault();
        setError('');
        setSuccessMsg('');
        setLoading(true);

        try {
            if (isLogin) {
                await auth.signInWithEmailAndPassword(email, password);
                // AuthContext akan otomatis mendeteksi auth state change
            } else {
                if (!name.trim()) throw new Error("Nama lengkap harus diisi");
                // Register
                const userCredential = await auth.createUserWithEmailAndPassword(email, password);
                
                // Set default user data in database
                await db.ref(`pmc_users/${userCredential.user.uid}`).set({
                    email: email,
                    name: name,
                    role: 'Guest', // Default role sebelum di-approve Admin
                    createdAt: new Date().toISOString()
                });
                
                setSuccessMsg("Pendaftaran berhasil! Akun Anda sedang menunggu persetujuan Super Admin.");
                setIsLogin(true); // Kembali ke form login
            }
        } catch (err) {
            let errorMsg = err.message;
            if (err.code === 'auth/invalid-credential' || err.code === 'auth/user-not-found' || err.code === 'auth/wrong-password') {
                errorMsg = 'Email atau kata sandi yang Anda masukkan salah. Silakan coba lagi.';
            } else if (err.code === 'auth/invalid-email') {
                errorMsg = 'Format email tidak valid.';
            } else if (err.code === 'auth/email-already-in-use') {
                errorMsg = 'Email ini sudah terdaftar. Silakan gunakan email lain atau masuk ke sistem.';
            } else if (err.code === 'auth/too-many-requests') {
                errorMsg = 'Akses ditolak karena terlalu banyak percobaan masuk yang gagal. Silakan coba lagi nanti.';
            } else if (err.code === 'auth/weak-password') {
                errorMsg = 'Kata sandi terlalu lemah. Gunakan minimal 6 karakter.';
            }
            setError(errorMsg);
        } finally {
            setLoading(false);
        }
    };

    return (
        <div className="min-h-screen bg-slate-900 flex flex-col md:flex-row font-sans text-slate-100 overflow-hidden relative">
            {/* Background Animations & Orbs */}
            <div className="absolute top-[-10%] left-[-10%] w-[500px] h-[500px] rounded-full bg-emerald-600/20 blur-[120px] pointer-events-none animate-pulse"></div>
            <div className="absolute bottom-[-10%] right-[-10%] w-[600px] h-[600px] rounded-full bg-blue-600/20 blur-[150px] pointer-events-none animate-pulse" style={{ animationDelay: '2s' }}></div>

            {/* Left Side: Branding / Visuals (Hidden on small screens) */}
            <div className="hidden md:flex md:w-1/2 lg:w-[55%] relative flex-col justify-between p-12 lg:p-20 border-r border-white/10 shadow-[20px_0_40px_rgba(0,0,0,0.3)] z-10">
                {/* Background Image with Overlay */}
                <div className="absolute inset-0 z-0">
                    <img 
                        src="https://images.unsplash.com/photo-1486406146926-c627a92ad1ab?q=80&w=2070&auto=format&fit=crop" 
                        alt="Modern Architecture" 
                        className="w-full h-full object-cover opacity-30 mix-blend-overlay"
                    />
                    <div className="absolute inset-0 bg-gradient-to-br from-slate-900/90 via-slate-900/70 to-emerald-900/40"></div>
                </div>

                <div className="relative z-10">
                    <div className="flex items-center gap-3 mb-12">
                        <img src={logoSidamon} alt="SIDAMON Logo" className="w-12 h-12 object-contain drop-shadow-md" />
                        <h1 className="text-3xl font-black tracking-tight text-white">SIDAMON<span className="text-emerald-500">.</span></h1>
                    </div>
                    
                    <h2 className="text-5xl lg:text-6xl font-bold leading-[1.1] tracking-tight text-white mb-6">
                        Enterprise <br /> 
                        <span className="text-transparent bg-clip-text bg-gradient-to-r from-emerald-400 to-cyan-400">
                            Resource Planning
                        </span>
                    </h2>
                    <p className="text-lg text-slate-300 max-w-md leading-relaxed border-l-4 border-emerald-500 pl-4">
                        Sistem Informasi Database & Monitoring Gaharu Sempana Group. Kelola proyek, inventaris, dan tenaga ahli dalam satu ekosistem cerdas.
                    </p>
                </div>


            </div>

            {/* Right Side: Auth Form */}
            <div className="w-full md:w-1/2 lg:w-[45%] flex items-center justify-center p-6 sm:p-12 relative z-10 min-h-screen md:min-h-0">
                <div className="w-full max-w-md">
                    {/* Mobile Branding (Visible only on small screens) */}
                    <div className="md:hidden text-center mb-10">
                        <img src={logoSidamon} alt="SIDAMON Logo" className="w-16 h-16 object-contain mx-auto drop-shadow-lg mb-6" />
                        <h1 className="text-3xl font-black text-white tracking-tight mb-2">SIDAMON<span className="text-emerald-500">.</span></h1>
                        <p className="text-sm text-slate-400">Gaharu Sempana Group</p>
                    </div>

                    <div className="bg-slate-800/40 backdrop-blur-2xl border border-white/10 p-8 sm:p-10 rounded-[2rem] shadow-2xl relative overflow-hidden">
                        {/* Decorative top gradient line */}
                        <div className="absolute top-0 left-0 w-full h-1 bg-gradient-to-r from-emerald-500 via-cyan-500 to-blue-500"></div>

                        <div className="mb-8">
                            <h2 className="text-2xl font-bold text-white mb-2">{isLogin ? 'Selamat Datang' : 'Buat Akun Baru'}</h2>
                            <p className="text-sm text-slate-400">
                                {isLogin ? 'Silahkan masukkan akun anda yang terdaftar untuk melanjutkan' : 'Daftarkan diri Anda untuk mengakses sistem.'}
                            </p>
                        </div>

                        {error && (
                            <div className="mb-6 p-4 rounded-xl bg-red-500/10 border border-red-500/20 flex items-start gap-3 fade-in">
                                <Icon name="alert-circle" size={18} className="text-red-400 shrink-0 mt-0.5" />
                                <span className="text-red-300 text-sm font-medium">{error}</span>
                            </div>
                        )}
                        
                        {successMsg && (
                            <div className="mb-6 p-4 rounded-xl bg-emerald-500/10 border border-emerald-500/20 flex items-start gap-3 fade-in">
                                <Icon name="check-circle-2" size={18} className="text-emerald-400 shrink-0 mt-0.5" />
                                <span className="text-emerald-300 text-sm font-medium">{successMsg}</span>
                            </div>
                        )}

                        <form onSubmit={handleSubmit} className="space-y-5">
                            {!isLogin && (
                                <div className="space-y-1.5">
                                    <label className="text-xs font-bold text-slate-400 uppercase tracking-wider ml-1">Nama Lengkap</label>
                                    <div className="relative group">
                                        <div className="absolute inset-y-0 left-0 pl-4 flex items-center pointer-events-none text-slate-500 group-focus-within:text-emerald-400 transition-colors">
                                            <Icon name="user" size={18} />
                                        </div>
                                        <input 
                                            type="text" 
                                            className="w-full bg-slate-900/50 border border-slate-700/50 rounded-2xl pl-11 pr-4 py-3.5 text-sm text-slate-100 focus:outline-none focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500 transition-all shadow-inner placeholder:text-slate-600"
                                            placeholder="John Doe"
                                            value={name}
                                            onChange={(e) => setName(e.target.value)}
                                            required={!isLogin}
                                        />
                                    </div>
                                </div>
                            )}
                            <div className="space-y-1.5">
                                <label className="text-xs font-bold text-slate-400 uppercase tracking-wider ml-1">Email</label>
                                <div className="relative group">
                                    <div className="absolute inset-y-0 left-0 pl-4 flex items-center pointer-events-none text-slate-500 group-focus-within:text-emerald-400 transition-colors">
                                        <Icon name="mail" size={18} />
                                    </div>
                                    <input 
                                        type="email" 
                                        className="w-full bg-slate-900/50 border border-slate-700/50 rounded-2xl pl-11 pr-4 py-3.5 text-sm text-slate-100 focus:outline-none focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500 transition-all shadow-inner placeholder:text-slate-600"
                                        placeholder="nama@email.com"
                                        value={email}
                                        onChange={(e) => setEmail(e.target.value)}
                                        required
                                    />
                                </div>
                            </div>
                            <div className="space-y-1.5">
                                <label className="text-xs font-bold text-slate-400 uppercase tracking-wider ml-1">Password</label>
                                <div className="relative group">
                                    <div className="absolute inset-y-0 left-0 pl-4 flex items-center pointer-events-none text-slate-500 group-focus-within:text-emerald-400 transition-colors">
                                        <Icon name="lock" size={18} />
                                    </div>
                                    <input 
                                        type="password" 
                                        className="w-full bg-slate-900/50 border border-slate-700/50 rounded-2xl pl-11 pr-4 py-3.5 text-sm text-slate-100 focus:outline-none focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500 transition-all shadow-inner placeholder:text-slate-600"
                                        placeholder="••••••••"
                                        value={password}
                                        onChange={(e) => setPassword(e.target.value)}
                                        required
                                    />
                                </div>
                            </div>

                            <button 
                                type="submit" 
                                disabled={loading}
                                className="w-full group relative overflow-hidden bg-emerald-600 hover:bg-emerald-500 text-white font-bold text-sm py-4 px-4 rounded-2xl transition-all shadow-lg shadow-emerald-900/40 disabled:opacity-70 disabled:cursor-not-allowed mt-4"
                            >
                                <div className="absolute inset-0 w-full h-full bg-gradient-to-r from-transparent via-white/20 to-transparent -translate-x-full group-hover:animate-[shimmer_1.5s_infinite]"></div>
                                <span className="relative flex items-center justify-center gap-2">
                                    {loading ? (
                                        <><Icon name="loader-2" size={18} className="animate-spin" /> Memproses...</>
                                    ) : (
                                        isLogin ? 'Masuk ke Sistem' : 'Daftar Sekarang'
                                    )}
                                </span>
                            </button>
                        </form>

                        <div className="mt-8 text-center border-t border-slate-700/50 pt-6">
                            <button 
                                type="button" 
                                onClick={() => { setIsLogin(!isLogin); setError(''); setSuccessMsg(''); }}
                                className="text-sm text-slate-400 hover:text-white transition-colors font-medium flex items-center justify-center gap-2 mx-auto"
                            >
                                {isLogin ? (
                                    <>Belum punya akun? <span className="text-emerald-400 font-bold">Daftar sekarang</span></>
                                ) : (
                                    <>Sudah punya akun? <span className="text-emerald-400 font-bold">Masuk di sini</span></>
                                )}
                            </button>
                        </div>
                    </div>
                    
                    <div className="mt-8 text-center md:hidden">
                        <p className="text-xs text-slate-500">&copy; 2026 Gaharu Sempana Group. All rights reserved.</p>
                    </div>
                </div>
            </div>
            
            {/* Styles for Custom Animations */}
            <style jsx>{`
                @keyframes shimmer {
                    100% { transform: translateX(100%); }
                }
            `}</style>
        </div>
    );
}
