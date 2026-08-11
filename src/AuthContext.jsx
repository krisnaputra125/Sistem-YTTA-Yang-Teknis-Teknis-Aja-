import React, { createContext, useContext, useState, useEffect } from 'react';
import { auth, db } from './firebase';

const AuthContext = createContext();

export const useAuth = () => useContext(AuthContext);

export const AuthProvider = ({ children }) => {
    const [currentUser, setCurrentUser] = useState(null);
    const [userRole, setUserRole] = useState(null); // String: "Super Admin", dll.
    const [loading, setLoading] = useState(true);

    useEffect(() => {
        const unsubscribe = auth.onAuthStateChanged(async (user) => {
            setCurrentUser(user);
            if (user) {
                const roleRef = db.ref(`pmc_users/${user.uid}`);
                roleRef.on('value', (snapshot) => {
                    const data = snapshot.val();
                    if (data && data.role) {
                        setUserRole(data.role);
                        setLoading(false);
                    } else {
                        // Auto-assign Super Admin to the first user
                        db.ref('pmc_users').once('value', usersSnapshot => {
                            if (!usersSnapshot.exists()) {
                                db.ref(`pmc_users/${user.uid}`).set({ role: 'Super Admin', email: user.email });
                                setUserRole('Super Admin');
                            } else {
                                db.ref(`pmc_users/${user.uid}`).set({ role: 'Guest', email: user.email });
                                setUserRole('Guest'); // Default role
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
                        alert("Sesi Anda telah berakhir karena tidak ada aktivitas selama 30 menit demi keamanan data. Silakan login kembali.");
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

        switch (menuName) {
            case 'Proyek':
            case 'Alokasi Tim':
            case 'Timesheet':
            case 'Time Schedule':
                // Menu Project Management & Control
                return ['Manajer Teknis', 'Kordinator Divisi Teknis', 'PIC', 'Team Leader Pekerjaan'].includes(userRole);
            
            case 'Tenaga Ahli':
            case 'Rekan Rekanan':
            case 'Manajemen LPSE':
                // Menu Database & Assignment Expert
                return ['Manajer Administrasi'].includes(userRole);

            case 'Manajemen Pengguna':
                return userRole === 'Super Admin';
                
            case 'Inventaris':
                return true;

            case 'Admin Aset':
                return ['Kordinator Aset'].includes(userRole);
                
            case 'KPI':
                return userRole === 'Super Admin';
                
            default:
                return false;
        }
    };

    const canCreateProject = () => {
        return ['Super Admin', 'Manajer Teknis', 'Kordinator Divisi Teknis'].includes(userRole);
    };

    const canDeleteProject = () => {
        return ['Super Admin', 'Manajer Teknis'].includes(userRole);
    };

    const canEditProjectAdmin = () => {
        return ['Super Admin', 'Manajer Administrasi'].includes(userRole);
    };

    const canEditProjectTechnical = () => {
        return ['Super Admin', 'Manajer Teknis', 'Kordinator Divisi Teknis'].includes(userRole);
    };

    const canEditTeamAllocation = () => {
        // PIC bisa edit alokasi tim, Manajer Teknis juga
        // Kordinator Divisi Teknis TIDAK BISA edit struktur alokasi tim
        return ['Super Admin', 'Manajer Teknis', 'PIC'].includes(userRole);
    };

    const canEditInventory = () => {
        return ['Super Admin', 'Kordinator Aset'].includes(userRole);
    };

    const canEditExperts = () => {
        return ['Super Admin', 'Manajer Administrasi'].includes(userRole);
    };

    const value = {
        currentUser,
        userRole,
        loading,
        canAccessMenu,
        canCreateProject,
        canDeleteProject,
        canEditProjectAdmin,
        canEditProjectTechnical,
        canEditTeamAllocation,
        canEditInventory,
        canEditExperts
    };

    return (
        <AuthContext.Provider value={value}>
            {!loading && children}
        </AuthContext.Provider>
    );
};
