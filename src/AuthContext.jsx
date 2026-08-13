import React, { createContext, useContext, useState, useEffect } from 'react';
import { auth, db } from './firebase';

const AuthContext = createContext();

export const useAuth = () => useContext(AuthContext);

export const AuthProvider = ({ children }) => {
    const [currentUser, setCurrentUser] = useState(null);
    const [userRole, setUserRole] = useState(null); // String: "Super Admin", dll.
    const [username, setUsername] = useState(null);
    const [loading, setLoading] = useState(true);

    useEffect(() => {
        const unsubscribe = auth.onAuthStateChanged(async (user) => {
            setCurrentUser(user);
            if (user) {
                const roleRef = db.ref(`pmc_users/${user.uid}`);
                roleRef.on('value', (snapshot) => {
                    const data = snapshot.val();
                    if (data && data.role) {
                        if (data.role === 'Deleted') {
                            auth.signOut().then(() => {
                                window.dispatchEvent(new CustomEvent('show-alert', { detail: { title: 'Akses Ditolak', message: 'Akun Anda telah dihapus secara permanen dari sistem.' } }));
                            });
                            setUserRole(null);
                            setLoading(false);
                            return;
                        }
                        setUserRole(data.role);
                        setUsername(data.username || data.name || null);
                        setLoading(false);
                    } else {
                        // Auto-assign Super Admin to the first user
                        db.ref('pmc_users').once('value', usersSnapshot => {
                            if (!usersSnapshot.exists()) {
                                db.ref(`pmc_users/${user.uid}`).set({ role: 'Super Admin', email: user.email });
                                setUserRole('Super Admin');
                            } else {
                                setUserRole('Guest'); // Default local state, do NOT write to DB so deleted users don't resurrect
                            }
                            setLoading(false);
                        });
                    }
                });
            } else {
                setUserRole(null);
                setLoading(false);
            }
        });

        return unsubscribe;
    }, []);

    // Fitur Auto Log Out setelah 30 menit tidak ada aktivitas (1800000 ms)
    useEffect(() => {
        let inactivityTimeout;

        const resetInactivityTimeout = () => {
            if (inactivityTimeout) clearTimeout(inactivityTimeout);
            inactivityTimeout = setTimeout(() => {
                if (auth.currentUser) {
                    auth.signOut().then(() => {
                        window.dispatchEvent(new CustomEvent('show-alert', { detail: { title: 'Sesi Berakhir', message: 'Sesi Anda telah berakhir karena tidak ada aktivitas selama 30 menit demi keamanan data. Silakan login kembali.' } }));
                    }).catch(console.error);
                }
            }, 1800000);
        };

        if (currentUser) {
            resetInactivityTimeout();
            const events = ['mousedown', 'mousemove', 'keypress', 'scroll', 'touchstart'];
            events.forEach(event => window.addEventListener(event, resetInactivityTimeout));

            return () => {
                if (inactivityTimeout) clearTimeout(inactivityTimeout);
                events.forEach(event => window.removeEventListener(event, resetInactivityTimeout));
            };
        }
    }, [currentUser]);

    // RBAC Permissions Logic
    const canAccessMenu = (menuName) => {
        if (!userRole || userRole === 'Guest') return false;
        if (userRole === 'Super Admin') return true;

        if (userRole === 'Manajer Teknis' || userRole === 'Manajer Administrasi' || userRole === 'HRD') {
            if (menuName === 'Manajemen Pengguna') return false;
            if (menuName === 'Admin Aset' && userRole !== 'HRD') return false;
            return true;
        }

        switch (menuName) {
            case 'Proyek':
                return ['Kordinator Divisi Teknis', 'PIC', 'Team Leader Pekerjaan', 'Manajer', 'Kordinator Tender'].includes(userRole);
            case 'Timesheet':
            case 'Time Schedule':
                return ['Kordinator Divisi Teknis', 'PIC', 'Team Leader Pekerjaan', 'Manajer'].includes(userRole);
                
            case 'Alokasi Tim':
            case 'Plotting Jadwal':
                return ['Kordinator Divisi Teknis', 'PIC', 'Manajer'].includes(userRole);
            
            case 'Tenaga Ahli':
            case 'Manajemen LPSE':
                return ['Manajer', 'Kordinator Tender'].includes(userRole);

            case 'Rekan Rekanan':
                return ['Manajer'].includes(userRole);

            case 'Manajemen Pengguna':
                return false;
                
            case 'Inventaris':
                return true;

            case 'Admin Aset':
                return ['Kordinator Aset', 'Manajer'].includes(userRole);
                
            case 'KPI':
                return ['Manajer'].includes(userRole);
                
            default:
                return false;
        }
    };

    const canCreateProject = () => {
        return ['Super Admin', 'Manajer Teknis', 'Kordinator Divisi Teknis', 'PIC', 'Team Leader Pekerjaan'].includes(userRole);
    };

    const canDeleteProject = () => {
        return ['Super Admin', 'Manajer Teknis', 'Kordinator Divisi Teknis', 'PIC', 'Team Leader Pekerjaan'].includes(userRole);
    };

    const canEditProjectAdmin = () => {
        return ['Super Admin', 'Manajer Administrasi'].includes(userRole);
    };

    const canEditProjectTechnical = () => {
        return ['Super Admin', 'Manajer Teknis', 'Kordinator Divisi Teknis', 'PIC', 'Team Leader Pekerjaan'].includes(userRole);
    };

    const canEditTeamAllocation = () => {
        return ['Super Admin', 'Manajer Teknis', 'Kordinator Divisi Teknis', 'PIC'].includes(userRole);
    };

    const canEditInventory = () => {
        return ['Super Admin', 'Kordinator Aset', 'HRD'].includes(userRole);
    };

    const canEditExperts = () => {
        return ['Super Admin', 'Manajer Administrasi', 'HRD', 'Kordinator Tender'].includes(userRole);
    };

    const canManageAssignments = () => {
        return ['Super Admin', 'Manajer Administrasi', 'HRD', 'Kordinator Tender'].includes(userRole);
    };

    const canManageAsset = () => {
        return ['Super Admin', 'Kordinator Aset', 'HRD'].includes(userRole);
    };

    const value = {
        currentUser,
        userRole,
        username,
        loading,
        canAccessMenu,
        canCreateProject,
        canDeleteProject,
        canEditProjectAdmin,
        canEditProjectTechnical,
        canEditTeamAllocation,
        canEditInventory,
        canEditExperts,
        canManageAssignments,
        canManageAsset
    };

    return (
        <AuthContext.Provider value={value}>
            {!loading && children}
        </AuthContext.Provider>
    );
};
