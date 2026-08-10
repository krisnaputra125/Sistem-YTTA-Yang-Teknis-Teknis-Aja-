import React, { useState } from 'react';
import { auth, db } from './firebase';

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
                
                setSuccessMsg("Pendaftaran berhasil! Akun Anda sedang menunggu persetujuan Super Admin (Role: Guest).");
                setIsLogin(true); // Kembali ke form login
            }
        } catch (err) {
            setError(err.message);
        } finally {
            setLoading(false);
        }
    };

    return (
        <div className="min-h-screen bg-slate-900 flex items-center justify-center p-4 relative overflow-hidden">
            {/* Background Decorations */}
            <div className="absolute top-[-20%] left-[-10%] w-[500px] h-[500px] rounded-full bg-emerald-600/20 blur-[120px] pointer-events-none"></div>
            <div className="absolute bottom-[-20%] right-[-10%] w-[600px] h-[600px] rounded-full bg-blue-600/20 blur-[150px] pointer-events-none"></div>

            <div className="w-full max-w-md bg-slate-800/80 backdrop-blur-xl border border-slate-700 p-8 rounded-3xl shadow-2xl z-10">
                <div className="text-center mb-8">
                    <h1 className="text-2xl font-black text-white tracking-tight mb-2">SIDAMON<span className="text-emerald-500">.</span></h1>
                    <p className="text-sm text-slate-400">Gaharu Sempana Group</p>
                </div>

                {error && (
                    <div className="mb-4 p-3 rounded-lg bg-red-900/30 border border-red-800 text-red-300 text-xs font-medium text-center">
                        {error}
                    </div>
                )}
                
                {successMsg && (
                    <div className="mb-4 p-3 rounded-lg bg-emerald-900/30 border border-emerald-800 text-emerald-300 text-xs font-medium text-center">
                        {successMsg}
                    </div>
                )}

                <form onSubmit={handleSubmit} className="space-y-4">
                    {!isLogin && (
                        <div>
                            <label className="block text-xs font-semibold text-slate-400 mb-1">Nama Lengkap</label>
                            <input 
                                type="text" 
                                className="w-full bg-slate-900/50 border border-slate-700 rounded-xl px-4 py-3 text-sm text-slate-200 focus:outline-none focus:border-emerald-500 transition-colors"
                                placeholder="Masukkan nama lengkap"
                                value={name}
                                onChange={(e) => setName(e.target.value)}
                                required={!isLogin}
                            />
                        </div>
                    )}
                    <div>
                        <label className="block text-xs font-semibold text-slate-400 mb-1">Email</label>
                        <input 
                            type="email" 
                            className="w-full bg-slate-900/50 border border-slate-700 rounded-xl px-4 py-3 text-sm text-slate-200 focus:outline-none focus:border-emerald-500 transition-colors"
                            placeholder="nama@email.com"
                            value={email}
                            onChange={(e) => setEmail(e.target.value)}
                            required
                        />
                    </div>
                    <div>
                        <label className="block text-xs font-semibold text-slate-400 mb-1">Password</label>
                        <input 
                            type="password" 
                            className="w-full bg-slate-900/50 border border-slate-700 rounded-xl px-4 py-3 text-sm text-slate-200 focus:outline-none focus:border-emerald-500 transition-colors"
                            placeholder="••••••••"
                            value={password}
                            onChange={(e) => setPassword(e.target.value)}
                            required
                        />
                    </div>

                    <button 
                        type="submit" 
                        disabled={loading}
                        className="w-full bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-sm py-3 px-4 rounded-xl transition-all shadow-lg shadow-emerald-900/50 disabled:opacity-50 disabled:cursor-not-allowed mt-2"
                    >
                        {loading ? 'Memproses...' : (isLogin ? 'Masuk ke Sistem' : 'Daftar Akun')}
                    </button>
                </form>

                <div className="mt-6 text-center">
                    <button 
                        type="button" 
                        onClick={() => { setIsLogin(!isLogin); setError(''); setSuccessMsg(''); }}
                        className="text-xs text-slate-400 hover:text-emerald-400 transition-colors font-medium"
                    >
                        {isLogin ? 'Belum punya akun? Daftar sekarang' : 'Sudah punya akun? Masuk di sini'}
                    </button>
                </div>
            </div>
        </div>
    );
}
