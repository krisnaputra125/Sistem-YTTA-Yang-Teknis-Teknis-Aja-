import React, { useState, useEffect, useRef, useMemo, useCallback } from 'react';
import logoImg from '../LGIHT TRANSPARAN (1).PNG';
import firebase from 'firebase/compat/app';
import 'firebase/compat/database';
import * as XLSX from 'xlsx-js-style';


        // --- SISTEM ICON INTERNAL ---
        const Icon = ({ name, size = 20, className = "" }) => {
            const paths = {
                "layout-dashboard": '<rect width="7" height="9" x="3" y="3" rx="1"/><rect width="7" height="5" x="14" y="3" rx="1"/><rect width="7" height="9" x="14" y="12" rx="1"/><rect width="7" height="5" x="3" y="16" rx="1"/>',
                "briefcase": '<rect width="20" height="14" x="2" y="7" rx="2" ry="2"/><path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"/>',
                "users": '<path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M22 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/>',
                "calendar": '<rect width="18" height="18" x="3" y="4" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/>',
                "chevron-down": '<polyline points="6 9 12 15 18 9"/>',
                "chevron-right": '<polyline points="9 18 15 12 9 6"/>',
                "calendar-days": '<rect width="18" height="18" x="3" y="4" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/><path d="M8 14h.01"/><path d="M12 14h.01"/><path d="M16 14h.01"/><path d="M8 18h.01"/><path d="M12 18h.01"/><path d="M16 18h.01"/>',
                "calendar-clock": '<path d="M21 7.5V6a2 2 0 0 0-2-2H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h3.5"/><path d="M16 2v4"/><path d="M8 2v4"/><path d="M3 10h5"/><path d="M17.5 17.5 16 16.25V14"/><circle cx="16" cy="16" r="6"/>',
                "search": '<circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3"/>',
                "database": '<ellipse cx="12" cy="5" rx="9" ry="3"/><path d="M3 5V19A9 3 0 0 0 21 19V5"/><path d="M3 12A9 3 0 0 0 21 12"/>',
                "refresh-ccw": '<path d="M21 12a9 9 0 0 0-9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/><path d="M3 3v5h5"/><path d="M3 12a9 9 0 0 0 9 9 9.75 9.75 0 0 0 6.74-2.74L21 16"/><path d="M16 21v-5h5"/>',
                "trending-up": '<polyline points="22 7 13.5 15.5 8.5 10.5 2 17"/><polyline points="16 7 22 7 22 13"/>',
                "trending-down": '<polyline points="22 17 13.5 8.5 8.5 13.5 2 7"/><polyline points="16 17 22 17 22 11"/>',
                "alert-triangle": '<path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3Z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>',
                "file-text": '<path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><line x1="10" y1="9" x2="8" y2="9"/>',
                "check-circle-2": '<path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/>',
                "award": '<circle cx="12" cy="8" r="6"/><path d="M15.477 12.89 17 22l-5-3-5 3 1.523-9.11"/>',
                "shield": '<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>',
                "edit-3": '<path d="M12 20h9"/><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"/>',
                "edit-2": '<path d="M17 3a2.828 2.828 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5L17 3z"/>',
                "phone": '<path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72 12.84 12.84 0 0 0 .7 2.81 2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45 12.84 12.84 0 0 0 2.81.7A2 2 0 0 1 22 16.92z"/>',
                "trash-2": '<polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/>',
                "clock": '<circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/>',
                "folder-open": '<path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/>',
                "plus": '<path d="M5 12h14"/><path d="M12 5v14"/>',
                "edit": '<path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>',
                "trash": '<path d="M3 6h18"/><path d="M19 6v14c0 1-1 2-2 2H7c-1 0-2-1-2-2V6"/><path d="M8 6V4c0-1 1-2 2-2h4c1 0 2 1 2 2v2"/>',
                "x": '<path d="M18 6 6 18"/><path d="M6 6l12 12"/>',
                "save": '<path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/>',
                "check": '<polyline points="20 6 9 17 4 12"/>',
                "target": '<circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="6"/><circle cx="12" cy="12" r="2"/>',
                "arrow-left": '<line x1="19" y1="12" x2="5" y2="12"/><polyline points="12 19 5 12 12 5"/>',
                "menu": '<line x1="4" y1="12" x2="20" y2="12"/><line x1="4" y1="6" x2="20" y2="6"/><line x1="4" y1="18" x2="20" y2="18"/>',
                "star": '<polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/>',
                "bar-chart": '<line x1="12" y1="20" x2="12" y2="10"/><line x1="18" y1="20" x2="18" y2="4"/><line x1="6" y1="20" x2="6" y2="16"/>',
                "help-circle": '<circle cx="12" cy="12" r="10"/><path d="M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3"/><line x1="12" y1="17" x2="12.01" y2="17"/>',
                "zap": '<polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/>',
                "sun": '<circle cx="12" cy="12" r="4"/><path d="M12 2v2"/><path d="M12 20v2"/><path d="m4.93 4.93 1.41 1.41"/><path d="m17.66 17.66 1.41 1.41"/><path d="M2 12h2"/><path d="M20 12h2"/><path d="m6.34 17.66-1.41 1.41"/><path d="m19.07 4.93-1.41 1.41"/>',
                "moon": '<path d="M12 3a6 6 0 0 0 9 9 9 9 0 1 1-9-9Z"/>',
                "hexagon": '<path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/>',
                "box": '<path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/><polyline points="3.27 6.96 12 12.01 20.73 6.96"/><line x1="12" y1="22.08" x2="12" y2="12"/>',
                "info": '<circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/>',
                "printer": '<polyline points="6 9 6 2 18 2 18 9"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect width="12" height="8" x="6" y="14"/>',
                "shopping-cart": '<circle cx="9" cy="21" r="1"/><circle cx="20" cy="21" r="1"/><path d="M1 1h4l2.68 13.39a2 2 0 0 0 2 1.61h9.72a2 2 0 0 0 2-1.61L23 6H6"/>',
                "play": '<polygon points="5 3 19 12 5 21 5 3"/>',
                "pie-chart": '<path d="M21.21 15.89A10 10 0 1 1 8 2.83"/><path d="M22 12A10 10 0 0 0 12 2v10z"/>',
                "pause": '<rect x="6" y="4" width="4" height="16"/><rect x="14" y="4" width="4" height="16"/>',
                "settings": '<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"/>',
                "link": '<path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/>'
            };
            return <svg xmlns="http://www.w3.org/2000/svg" width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className} dangerouslySetInnerHTML={{ __html: paths[name] || '' }} />;
        };

        // --- HELPER FUNCTION: FORMAT TANGGAL & HITUNG STATUS ---
        const formatDateTimeIndo = (dateString) => {
            if (!dateString) return "-";
            try {
                const date = new Date(dateString);
                if (isNaN(date.getTime())) return dateString;

                const months = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agt', 'Sep', 'Okt', 'Nov', 'Des'];
                const d = date.getDate();
                const m = months[date.getMonth()];
                const y = date.getFullYear();
                const h = String(date.getHours()).padStart(2, '0');
                const min = String(date.getMinutes()).padStart(2, '0');

                return `${d} ${m} ${y}, ${h}:${min}`;
            } catch (e) {
                return dateString;
            }
        };

        const formatDateIndo = (dateString) => {
            if (!dateString) return "-";
            try {
                const date = new Date(dateString);
                if (isNaN(date.getTime())) return dateString;

                const months = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
                const day = date.getDate();
                const month = months[date.getMonth()];
                const year = date.getFullYear();
                return `${day} ${month} ${year}`;
            } catch (e) {
                return dateString;
            }
        };

        const calculateComputedStatus = (project) => {
            if (project.isPending) return "Pending";

            const baseStatus = project.status === "On Track" ? "On Progress" : (project.status || "On Progress");
            if (baseStatus === "Done" || project.progress >= 100) return "Done";

            const today = new Date();
            today.setHours(0, 0, 0, 0);

            // LOGIC NOT STARTED UNTUK PENGAWASAN
            const isPengawasan = project.type?.toLowerCase().includes('pengawas') || project.type?.toLowerCase().includes('manajemen konstruksi');
            if (isPengawasan) {
                if (project.spmk) {
                    const spmkDate = new Date(project.spmk);
                    spmkDate.setHours(0, 0, 0, 0);
                    if (spmkDate > today) return "Not Started";
                } else {
                    return "Not Started";
                }
            }

            if (!project.deadline) return baseStatus;

            const deadlineDate = new Date(project.deadline);
            deadlineDate.setHours(0, 0, 0, 0);

            if (isNaN(deadlineDate.getTime())) return baseStatus;

            const diffTime = deadlineDate.getTime() - today.getTime();
            const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

            if (diffDays < 0) {
                if (isPengawasan) return "Done";
                return "Terlambat";
            } else if (diffDays <= 14) {
                if (isPengawasan) return "On Progress";
                return "Beresiko";
            }
            return baseStatus;
        };

        const getMicroStatus = (progress, deadlineStr) => {
            if (progress >= 100) return "Done";
            if (!deadlineStr) return "Belum Diatur";

            const today = new Date();
            today.setHours(0, 0, 0, 0);
            const dDate = new Date(deadlineStr);
            dDate.setHours(0, 0, 0, 0);

            if (isNaN(dDate.getTime())) return "Belum Diatur";

            const diffTime = dDate.getTime() - today.getTime();
            const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

            if (diffDays < 0) return "Terlambat";
            if (diffDays <= 3) return "Beresiko"; // Risiko Mikro H-3
            return "On Progress";
        };

        const getLPSEHierarchyScore = (roleStr) => {
            if (!roleStr) return 99;
            const r = roleStr.toLowerCase();
            if (r.includes('team leader') || r.includes('ketua tim')) return 1;
            if (r.includes('ahli utama')) return 2;
            if (r.includes('ahli madya')) return 3;
            if (r.includes('ahli muda')) return 4;
            if (r.includes('ahli') || r.includes('expert') || r.includes('spesialis')) return 5;
            if (r.includes('inspector') || r.includes('inspektur') || r.includes('pengawas')) return 6;
            if (r.includes('asisten') || r.includes('sub profesional') || r.includes('sub-profesional')) return 7;
            if (r.includes('surveyor') || r.includes('survey')) return 8;
            if (r.includes('drafter') || r.includes('cad') || r.includes('juru gambar')) return 9;
            if (r.includes('pendukung') || r.includes('support')) return 10;
            if (r.includes('administrasi') || r.includes('admin') || r.includes('sekretaris') || r.includes('operator')) return 1000;
            return 50; 
        };

        const getCategoryFromRole = (roleStr) => {
            if (!roleStr) return 'Lainnya';
            const roleLower = roleStr.toLowerCase();
            if (roleLower.includes('arsitek')) return 'Arsitek';
            if (roleLower === 'qs' || roleLower.includes('quantity')) return 'QS';
            if (roleLower.includes('struktur')) return 'Struktur';
            if (roleLower.includes('mep')) return 'MEP';
            if (roleLower.includes('tata ruang') || roleLower.includes('planologi')) return 'Tata Ruang';
            return 'Lainnya';
        };

        const getEffectiveEmpCategory = (project, employeeName, employeeRole) => {
            if (project.type?.toLowerCase().includes('perencana') && project.surveyorTeam?.includes(employeeName)) {
                return 'Surveyor';
            }
            return getCategoryFromRole(employeeRole);
        };

        const calculateLeaderKPI = (employee, allProjects) => {
            const isKordinator = employee.level?.startsWith('Kordinator Divisi');
            const kordinatorDivisi = isKordinator ? employee.level.replace('Kordinator Divisi ', 'Divisi ') : null;

            const ledProjects = allProjects.filter(p => p.teamLeader === employee.name);
            const subProjects = allProjects.filter(p => p.team && p.team.includes(employee.name));
            const kontrolProjects = isKordinator ? allProjects.filter(p => p.type?.toLowerCase().includes('perencana') && p.divisiKontrol === kordinatorDivisi) : [];

            // Gabungkan proyek dan hapus duplikat
            const involvedProjectsMap = new Map();
            ledProjects.forEach(p => involvedProjectsMap.set(p.id, p));
            subProjects.forEach(p => involvedProjectsMap.set(p.id, p));
            kontrolProjects.forEach(p => involvedProjectsMap.set(p.id, p));
            const allInvolvedProjects = Array.from(involvedProjectsMap.values());

            const involvedProjects = allInvolvedProjects.filter(p => {
                const isPengawasan = p.type?.toLowerCase().includes('pengawas') || p.type?.toLowerCase().includes('manajemen konstruksi');
                if (isPengawasan) {
                    const statusTurun = p.pengawasanDetails?.[employee.name]?.statusTurun;
                    if (statusTurun === 'Tidak Turun') return false;
                }
                return true;
            });

            const activeProjects = involvedProjects.filter(p => p.status !== "Done" && p.computedStatus !== "Done" && p.computedStatus !== "Pending");
            const totalActive = activeProjects.length;

            let score = 100;
            let delayedCount = 0;
            let atRiskCount = 0;
            let totalProgress = 0;
            let doneOnTimeBonus = 0;
            let staffDoneBonus = 0;

            involvedProjects.forEach(p => {
                const isLeader = p.teamLeader === employee.name || (isKordinator && p.divisiKontrol === kordinatorDivisi);
                const empCat = getEffectiveEmpCategory(p, employee.name, employee.role);

                if (isLeader) {
                    if (p.computedStatus === "Terlambat" && p.status !== "Done") {
                        delayedCount++;
                        score -= 10;
                    } else if (p.computedStatus === "Beresiko" && p.status !== "Done") {
                        atRiskCount++;
                    }

                    if (p.status === "Done" && p.computedStatus !== "Terlambat") {
                        doneOnTimeBonus++;
                        score += 20;
                    }

                    if (p.status !== "Done" && p.computedStatus !== "Done") {
                        totalProgress += Number(p.progress || 0);
                    }
                } else {
                    // STAFF LOGIC (MICRO)
                    let pProgress = 0;
                    let pStatus = "On Progress";
                    const microProgress = p.categoryDetails?.[empCat]?.progress ? Number(p.categoryDetails[empCat].progress) : 0;
                    const individualDone = (p.individualStatus && p.individualStatus[employee.name] === true) || microProgress === 100;

                    if (individualDone) {
                        pProgress = 100;
                        pStatus = "Done";
                    } else if (p.categoryDetails && p.categoryDetails[empCat]) {
                        const micro = p.categoryDetails[empCat];
                        pProgress = microProgress;
                        pStatus = getMicroStatus(pProgress, micro.deadline);
                    } else {
                        pProgress = Number(p.progress || 0);
                        pStatus = p.computedStatus;
                    }

                    if (pStatus === "Terlambat" && p.status !== "Done") {
                        delayedCount++;
                        score -= 5;
                    } else if (pStatus === "Beresiko" && p.status !== "Done") {
                        atRiskCount++;
                    }

                    if (p.status === "Done" && p.computedStatus !== "Terlambat") {
                        staffDoneBonus++;
                        score += 15;
                    }

                    if (p.status !== "Done" && p.computedStatus !== "Done") {
                        totalProgress += pProgress;
                    }
                }
            });

            const avgProgress = totalActive > 0 ? totalProgress / totalActive : 0;
            const isOverloaded = (totalActive * 25) > 100;

            if (isOverloaded) {
                if (delayedCount === 0) score += ((totalActive - 4) * 10);
                else score -= 10;
            }
            score += (employee.manualPoints || 0);

            if (score > 100) score = 100;
            if (score < 0) score = 0;

            let kpiStatus = "Sangat Baik";
            if (score < 60) kpiStatus = "Perlu Perhatian";
            else if (score < 80) kpiStatus = "Cukup";

            return {
                score: (totalActive === 0 && doneOnTimeBonus === 0 && staffDoneBonus === 0 && !(employee.manualPoints)) ? 100 : Math.round(score),
                delayed: delayedCount,
                atRisk: atRiskCount,
                avgProgress: Math.round(avgProgress),
                status: kpiStatus,
                isOverloaded,
                totalProjects: totalActive,
                rating: employee.rating || 3,
                bonusDone: doneOnTimeBonus,
                bonusDoneLeader: doneOnTimeBonus,
                bonusDoneStaff: staffDoneBonus
            };
        };

        const calculateEmployeeKPI = (employee, allProjects) => {
            const allInvolvedProjects = allProjects.filter(p => {
                const members = Array.isArray(p.team) ? p.team : [];
                return members.includes(employee.name);
            });

            const involvedProjects = allInvolvedProjects.filter(p => {
                const isPengawasan = p.type?.toLowerCase().includes('pengawas') || p.type?.toLowerCase().includes('manajemen konstruksi');
                if (isPengawasan) {
                    const statusTurun = p.pengawasanDetails?.[employee.name]?.statusTurun;
                    if (statusTurun === 'Tidak Turun') return false;
                }
                return true;
            });

            const activeProjects = involvedProjects.filter(p => p.status !== "Done" && p.computedStatus !== "Done" && p.computedStatus !== "Pending");

            const totalProjects = activeProjects.length;

            let delayedCount = 0;
            let atRiskCount = 0;
            let totalProgress = 0;
            let doneOnTimeBonus = 0;
            let score = 100;

            involvedProjects.forEach(p => {
                const empCat = getEffectiveEmpCategory(p, employee.name, employee.role);
                let pProgress = 0;
                let pStatus = "On Progress";
                const microProgress = p.categoryDetails?.[empCat]?.progress ? Number(p.categoryDetails[empCat].progress) : 0;
                const individualDone = (p.individualStatus && p.individualStatus[employee.name] === true) || microProgress === 100;

                if (individualDone) {
                    pProgress = 100;
                    pStatus = "Done";
                } else if (p.categoryDetails && p.categoryDetails[empCat]) {
                    const micro = p.categoryDetails[empCat];
                    pProgress = microProgress;
                    pStatus = getMicroStatus(pProgress, micro.deadline);
                } else {
                    pProgress = Number(p.progress || 0);
                    pStatus = p.computedStatus;
                }

                if (p.status === "Done" && p.computedStatus !== "Terlambat") {
                    doneOnTimeBonus++;
                    score += 15;
                }

                if (p.status !== "Done" && p.computedStatus !== "Done") {
                    totalProgress += pProgress;
                }

                if (pStatus === "Terlambat" && p.status !== "Done") delayedCount++;
                else if (pStatus === "Beresiko" && p.status !== "Done") atRiskCount++;
            });

            const avgProgress = totalProjects > 0 ? (totalProgress / totalProjects) : 0;
            const isOverloaded = (totalProjects * 25) > 100;

            score -= (delayedCount * 5);

            if (isOverloaded) {
                if (delayedCount === 0) {
                    score += ((totalProjects - 4) * 10);
                } else {
                    score -= 10;
                }
            }
            score += (employee.manualPoints || 0);

            if (score > 100) score = 100;
            if (score < 0) score = 0;

            let kpiStatus = "Sangat Baik";
            if (score < 60) kpiStatus = "Perlu Perhatian";
            else if (score < 80) kpiStatus = "Cukup";

            return {
                score: (totalProjects === 0 && doneOnTimeBonus === 0 && !(employee.manualPoints)) ? 100 : Math.round(score),
                delayed: delayedCount,
                atRisk: atRiskCount,
                avgProgress: Math.round(avgProgress),
                status: kpiStatus,
                isOverloaded,
                totalProjects,
                rating: employee.rating || 3,
                bonusDone: doneOnTimeBonus
            };
        };

        // --- KOMPONEN CHECKBOX DIPISAH AGAR SCROLL TIDAK RESET ---
        function TeamCheckboxGroup({ title, roleFilter, isOptional, filteredResources, formData, setFormData, teamField = 'team' }) {
            const groupResources = filteredResources.filter(roleFilter);
            if (groupResources.length === 0) return null;

            return (
                <div className="mb-3">
                    <h4 className="text-[11px] font-bold text-slate-700 bg-slate-200 px-2 py-1 rounded mb-1.5 flex justify-between items-center">
                        <span>{title}</span>
                        {isOptional && <span className="font-normal italic text-slate-500"></span>}
                    </h4>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-1">
                        {groupResources.map(res => (
                            <label key={res.id} className="flex items-start space-x-2 text-sm cursor-pointer p-1 hover:bg-slate-100 rounded transition-colors">
                                <input
                                    type="checkbox"
                                    className="mt-0.5 rounded border-slate-300 w-4 h-4 text-blue-600 focus:ring-blue-500"
                                    checked={(formData[teamField] || []).includes(res.name)}
                                    onChange={(e) => {
                                        const newTeam = e.target.checked
                                            ? [...(formData[teamField] || []), res.name]
                                            : (formData[teamField] || []).filter(n => n !== res.name);
                                        const updates = { [teamField]: newTeam };

                                        // Jika dicentang di sub-tim biasa, hapus dari surveyorTeam (mencegah centang ganda)
                                        if (e.target.checked && teamField === 'team' && formData.surveyorTeam && formData.surveyorTeam.includes(res.name)) {
                                            updates.surveyorTeam = formData.surveyorTeam.filter(n => n !== res.name);
                                        }

                                        setFormData({ ...formData, ...updates });
                                    }}
                                />
                                <div className="flex flex-col">
                                    <span className="text-[12px] font-medium leading-tight">{res.name}</span>
                                </div>
                            </label>
                        ))}
                    </div>
                </div>
            );
        }

        function TeamVisionaryMultiSelect({ title, filteredResources, formData, setFormData, teamField = 'surveyorTeam' }) {
            const [searchTerm, setSearchTerm] = React.useState('');

            const searchResults = filteredResources.filter(r =>
                searchTerm === '' ||
                r.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
                r.role.toLowerCase().includes(searchTerm.toLowerCase())
            );

            const togglePerson = (name) => {
                const current = formData[teamField] || [];
                if (current.includes(name)) {
                    setFormData({ ...formData, [teamField]: current.filter(n => n !== name) });
                } else {
                    const updates = { [teamField]: [...current, name] };

                    // Jika diplot ke surveyorTeam, hapus dari sub-tim biasa (team) agar tidak tercentang ganda
                    if (teamField === 'surveyorTeam' && formData.team && formData.team.includes(name)) {
                        updates.team = formData.team.filter(n => n !== name);
                    }

                    setFormData({ ...formData, ...updates });
                }
            };

            return (
                <div className="mb-4 bg-blue-50/40 dark:bg-blue-900/10 border border-blue-200 dark:border-blue-800/60 p-3 rounded-xl shadow-sm">
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-3">
                        <h4 className="text-[12px] font-bold text-blue-800 dark:text-blue-300 flex items-center gap-1.5">
                            <Icon name="users" size={14} className="text-blue-600 dark:text-blue-400" />
                            {title}
                        </h4>
                        <div className="relative w-full sm:w-1/2">
                            <div className="absolute left-2.5 top-1/2 -translate-y-1/2 text-blue-400">
                                <Icon name="search" size={12} />
                            </div>
                            <input
                                type="text"
                                placeholder="Cari nama/jabatan (cth: arsitek)..."
                                className="w-full pl-7 pr-3 py-1.5 bg-white dark:bg-slate-900 border border-blue-200 dark:border-blue-700/50 rounded-lg text-[11px] outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all shadow-sm dark:text-slate-200"
                                value={searchTerm}
                                onChange={(e) => setSearchTerm(e.target.value)}
                            />
                        </div>
                    </div>

                    {(formData[teamField] || []).length > 0 && (
                        <div className="flex flex-wrap gap-1.5 mb-3 p-2 bg-white dark:bg-slate-800/80 rounded-lg border border-blue-100 dark:border-blue-800/50 shadow-inner min-h-[36px] items-center">
                            <span className="text-[9px] font-bold text-slate-400 mr-1 uppercase tracking-wider">Terpilih:</span>
                            {(formData[teamField] || []).map(name => (
                                <div key={name} className="flex items-center gap-1 bg-blue-600 text-white px-2 py-0.5 rounded text-[10px] font-bold shadow-sm group">
                                    <span>{name}</span>
                                    <button
                                        type="button"
                                        onClick={() => togglePerson(name)}
                                        className="text-blue-200 hover:text-white bg-blue-700 hover:bg-red-500 rounded-full w-3.5 h-3.5 flex items-center justify-center transition-colors"
                                        title="Hapus"
                                    >
                                        ×
                                    </button>
                                </div>
                            ))}
                        </div>
                    )}

                    <div className="max-h-32 overflow-y-auto pr-1.5 space-y-1">
                        {searchResults.length === 0 ? (
                            <p className="text-[11px] text-slate-400 italic text-center py-3">Tidak ditemukan personil dengan pencarian "{searchTerm}".</p>
                        ) : (
                            <div className="grid grid-cols-1 md:grid-cols-2 gap-1.5">
                                {searchResults.map(res => {
                                    const isSelected = (formData[teamField] || []).includes(res.name);
                                    return (
                                        <label
                                            key={res.id}
                                            className={`flex items-start space-x-2 text-sm cursor-pointer p-1.5 rounded-lg transition-all border ${isSelected ? 'bg-blue-50/80 dark:bg-blue-900/40 border-blue-300 dark:border-blue-600 shadow-sm' : 'hover:bg-slate-100 dark:hover:bg-slate-800/50 border-transparent hover:border-slate-200 dark:hover:border-slate-700'}`}
                                        >
                                            <input
                                                type="checkbox"
                                                className="mt-0.5 rounded border-slate-300 dark:border-slate-600 w-3.5 h-3.5 text-blue-600 focus:ring-blue-500 transition-colors"
                                                checked={isSelected}
                                                onChange={() => togglePerson(res.name)}
                                            />
                                            <div className="flex flex-col">
                                                <span className={`text-[11px] font-bold leading-tight ${isSelected ? 'text-blue-800 dark:text-blue-300' : 'text-slate-700 dark:text-slate-200'}`}>{res.name}</span>
                                                <span className="text-[9px] text-slate-500 dark:text-slate-400 mt-0.5">{res.role}</span>
                                            </div>
                                        </label>
                                    )
                                })}
                            </div>
                        )}
                    </div>
                </div>
            );
        }

        let expertSaveTimeout = null;
        let assignmentSaveTimeout = null;

        function App() {
            const [isDarkMode, setIsDarkMode] = React.useState(() => {
                if (typeof window !== 'undefined') {
                    return localStorage.getItem('theme') === 'dark';
                }
                return false;
            });

            React.useEffect(() => {
                const root = document.documentElement;
                if (isDarkMode) {
                    root.classList.add('dark');
                    localStorage.setItem('theme', 'dark');
                } else {
                    root.classList.remove('dark');
                    localStorage.setItem('theme', 'light');
                }
            }, [isDarkMode]);

            const [sidebarOpen, setSidebarOpen] = useState(false);
            const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
            const [activeTab, setActiveTab] = useState('dashboard');
            const [activeScheduleProject, setActiveScheduleProject] = useState(null);
            const [scheduleZoom, setScheduleZoom] = useState('month');
            const [scheduleFilterUser, setScheduleFilterUser] = useState('Semua');
            const [scheduleCollapsedCats, setScheduleCollapsedCats] = useState({});

            const [chartAnimate, setChartAnimate] = useState(false);

            useEffect(() => {
                if (activeScheduleProject) {
                    setChartAnimate(false);
                    setTimeout(() => setChartAnimate(true), 100);
                }
            }, [activeScheduleProject]);
            const [scheduleViewType, setScheduleViewType] = useState('timeline');
            const [projects, setProjects] = useState([]);
            const [resources, setResources] = useState([]);
            const [inventory, setInventory] = useState([]);
            const [experts, setExperts] = useState([]);
            const [borrowCart, setBorrowCart] = useState([]);
            const [loading, setLoading] = useState(true);
            const [errorMsg, setErrorMsg] = useState("");
            const [isLive, setIsLive] = useState(false);

            const [printData, setPrintData] = useState(null);
            const [printZoomProject, setPrintZoomProject] = useState(null);
            const [showPrintModal, setShowPrintModal] = useState(false);
            const [printOptions, setPrintOptions] = useState({ projectType: 'Semua', section: 'Semua' });

            // Pending Modal States
            const [showPendingModal, setShowPendingModal] = useState(false);
            const [pendingProjectData, setPendingProjectData] = useState(null);
            const [pendingReasonText, setPendingReasonText] = useState("");

            // Resume Modal States
            const [showResumeModal, setShowResumeModal] = useState(false);
            const [resumeProjectData, setResumeProjectData] = useState(null);

            // Alert Modal States
            const [alertModal, setAlertModal] = useState({ isOpen: false, title: 'Perhatian', message: '' });
            React.useEffect(() => {
                // Print mode is now interactive. User must click "Simpan PDF / Cetak" manually.
            }, [printData]);

            // Firebase Configuration
            const firebaseConfig = {
                apiKey: "AIzaSyDerRfccB2e-k3SSLoo5gxh10sRy9FRRwY",
                authDomain: "project-management-control.firebaseapp.com",
                databaseURL: "https://project-management-control-default-rtdb.asia-southeast1.firebasedatabase.app/",
                projectId: "project-management-control",
                storageBucket: "project-management-control.firebasestorage.app",
                messagingSenderId: "1065130749964",
                appId: "1:1065130749964:web:404873304275127c85fe29",
                measurementId: "G-LLNQ6P9DQ9"
            };

            if (typeof firebase !== 'undefined' && !firebase.apps.length) {
                firebase.initializeApp(firebaseConfig);
            }
            const firebaseDbRef = typeof firebase !== 'undefined' ? firebase.database().ref('pmc_data') : null;

            // Fitur Sanitasi Firebase: Mengubah karakter dilarang seperti titik (.) menjadi _DOT_ sebelum disimpan, 
            // lalu mengembalikannya menjadi titik (.) saat dibaca agar nama tetap utuh di UI
            const encodeKey = (k) => typeof k === 'string' ? k.replace(/\./g, '__DOT__').replace(/#/g, '__HASH__').replace(/\$/g, '__DOLLAR__').replace(/\[/g, '__LBRACK__').replace(/\]/g, '__RBRACK__').replace(/\//g, '__SLASH__') : k;
            const decodeKey = (k) => typeof k === 'string' ? k.replace(/__DOT__/g, '.').replace(/__HASH__/g, '#').replace(/__DOLLAR__/g, '$').replace(/__LBRACK__/g, '[').replace(/__RBRACK__/g, ']').replace(/__SLASH__/g, '/') : k;

            // Old URL Web App Google Apps Script (Digunakan sekali saat migrasi otomatis)
            const GOOGLE_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyEt5puRB26FO6mGfnxVvofBJaDwLtH3O0yd-Ugsugn2D2KezKpi_ynGz7hw24kpLIpRQ/exec';
            const [modalConfig, setModalConfig] = useState({ isOpen: false, type: null, mode: 'add', data: null });
            const [confirmDialog, setConfirmDialog] = useState({ isOpen: false, title: '', message: '', type: 'info', onConfirm: null });
            const [lpseList, setLpseList] = useState([]);
            const [showLpseManager, setShowLpseManager] = useState(false);
            const [assignments, setAssignments] = useState([]);

            const handleUpdateLpseList = (newList) => {
                firebase.database().ref('pmc_lpse_list').set(newList);
                setLpseList(newList);
            };

            const [certList, setCertList] = useState([]);
            const [showCertManager, setShowCertManager] = useState(false);
            const handleUpdateCertList = (newList) => {
                firebase.database().ref('pmc_cert_list').set(newList);
                setCertList(newList);
            };

            // State Developer Auth (BETA)
            const [showDevPrompt, setShowDevPrompt] = useState(false);
            const [devPassword, setDevPassword] = useState("");
            const [devAuthError, setDevAuthError] = useState("");
            const [isDeveloper, setIsDeveloper] = useState(false);
            const [pendingTabAttempt, setPendingTabAttempt] = useState(null);
            const [showKPIInfoModal, setShowKPIInfoModal] = useState(false);
            const [showProjectTypeModal, setShowProjectTypeModal] = useState(false);

            // State pencarian
            const [searchTeamTab, setSearchTeamTab] = useState("");
            const [searchProjectTab, setSearchProjectTab] = useState("");
            const [filterProjectType, setFilterProjectType] = useState("Semua Tipe");
            const [searchKPITab, setSearchKPITab] = useState("");
            const [searchGanttTab, setSearchGanttTab] = useState("");
            const [searchInvTab, setSearchInvTab] = useState("");
            const [searchExpertTab, setSearchExpertTab] = useState("");
            const [expertPage, setExpertPage] = useState(1);
            const [searchAssignmentTab, setSearchAssignmentTab] = useState("");
            const [projectPage, setProjectPage] = useState(1);
            const [dominoAnalysis, setDominoAnalysis] = useState(null);

            const projectsPerPage = 5;

            // State untuk melihat detail pegawai
            const [viewingEmployee, setViewingEmployee] = useState(null);

            // =========================================================================
            // KECERDASAN SISTEM
            // =========================================================================
            const computedProjects = useMemo(() => {
                return projects.map(p => {
                    let effectiveCatDetails = { ...p.categoryDetails };
                    const individualStatus = p.individualStatus || {};
                    let activeCategories = [];

                    if (p.team && p.team.length > 0) {
                        const catMembers = {};
                        p.team.forEach(m => {
                            const res = resources.find(r => r.name === m);
                            const cat = res ? getCategoryFromRole(res.role) : 'Lainnya';
                            if (!catMembers[cat]) catMembers[cat] = [];
                            catMembers[cat].push(m);
                        });

                        if (p.surveyorTeam && p.surveyorTeam.length > 0) {
                            catMembers['Surveyor'] = p.surveyorTeam;
                        }

                        activeCategories = Object.keys(catMembers);
                        activeCategories.forEach(cat => {
                            const members = catMembers[cat];
                            const total = members.length;
                            const completedCount = members.filter(m => individualStatus[m]).length;
                            const notCompletedCount = total - completedCount;

                            const manualProgress = effectiveCatDetails[cat] ? Number(effectiveCatDetails[cat].progress || 0) : 0;

                            if (total > 0 && completedCount > 0) {
                                const effectiveProg = Math.round(((completedCount * 100) + (notCompletedCount * manualProgress)) / total);
                                if (!effectiveCatDetails[cat]) effectiveCatDetails[cat] = {};
                                effectiveCatDetails[cat] = { ...effectiveCatDetails[cat], progress: effectiveProg };
                            }
                        });
                    }

                    let effectiveTotalProgress = Number(p.progress || 0);
                    if (activeCategories.length > 0) {
                        let sum = 0;
                        activeCategories.forEach(cat => {
                            sum += Number(effectiveCatDetails[cat]?.progress || 0);
                        });
                        effectiveTotalProgress = Math.round(sum / activeCategories.length);
                    }

                    const newP = {
                        ...p,
                        categoryDetails: effectiveCatDetails,
                        progress: effectiveTotalProgress
                    };
                    newP.computedStatus = calculateComputedStatus(newP);
                    return newP;
                });
            }, [projects, resources]);

            const calculatedResources = useMemo(() => {
                const mapped = resources.map(res => {
                    let numProjects = 0;
                    let numActiveProjects = 0;
                    const empCat = getCategoryFromRole(res.role);

                    computedProjects.forEach(p => {
                        if (p.notStarted) return;
                        if (p.computedStatus !== 'Done') {
                            const isPengawasan = p.type?.toLowerCase().includes('pengawas') || p.type?.toLowerCase().includes('manajemen konstruksi');

                            // Jika pengawasan dan statusnya tidak turun, lewati (tidak dihitung beban)
                            if (isPengawasan) {
                                const statusTurun = p.pengawasanDetails?.[res.name]?.statusTurun;
                                if (statusTurun === 'Tidak Turun') return;
                            }

                            if (p.teamLeader === res.name) {
                                numProjects++;
                                if (p.computedStatus !== 'Pending') numActiveProjects++;
                            } else if (p.team && p.team.includes(res.name)) {
                                let isIndividuallyDone = p.individualStatus?.[res.name] === true;

                                if (isPengawasan) {
                                    const pengawasanDeadline = p.pengawasanDetails?.[res.name]?.deadline;
                                    if (pengawasanDeadline && pengawasanDeadline < new Date().toISOString().split('T')[0]) {
                                        isIndividuallyDone = true;
                                    }
                                    if (!isIndividuallyDone) {
                                        numProjects++;
                                        if (p.computedStatus !== 'Pending') numActiveProjects++;
                                    }
                                } else {
                                    const effectiveCat = getEffectiveEmpCategory(p, res.name, res.role);
                                    const microProgress = p.categoryDetails?.[effectiveCat]?.progress ? Number(p.categoryDetails[effectiveCat].progress) : 0;
                                    if (microProgress < 100 && !isIndividuallyDone) {
                                        numProjects++;
                                        if (p.computedStatus !== 'Pending') numActiveProjects++;
                                    }
                                }
                            }
                        }
                    });
                    const workload = numActiveProjects * 25; // 25% per active project (pending not included)
                    return { ...res, projects: numProjects, workload };
                });

                return mapped.sort((a, b) => b.workload - a.workload);
            }, [computedProjects, resources]);

            const problematicProjectsCount = computedProjects.filter(p => p.computedStatus === 'Terlambat' || p.computedStatus === 'Beresiko').length;
            const completedProjectsCount = computedProjects.filter(p => p.computedStatus === 'Done').length;


            // FETCH DATA FROM FIREBASE
            const initFirebaseListener = () => {
                if (!firebaseDbRef) return;
                setLoading(true);
                setErrorMsg("");

                firebaseDbRef.once('value', async (snapshot) => {
                    if (!snapshot.exists()) {
                        console.log("Firebase kosong. Memulai migrasi otomatis dari Google Sheets...");
                        try {
                            const response = await fetch(GOOGLE_SCRIPT_URL);
                            const data = await response.json();

                            // Transformasi awal saat migrasi
                            const initialData = {
                                projects: data.projects || [],
                                resources: data.resources || [],
                                inventory: data.inventory || []
                            };

                            await firebaseDbRef.set(initialData);
                            console.log("Migrasi selesai!");
                        } catch (error) {
                            console.error("Gagal migrasi:", error);
                        }
                    }

                    // Pasang Listener Khusus untuk Inventaris (dipisah dari pmc_data agar tidak tertimpa sync)
                    firebase.database().ref('pmc_inventory').on('value', (snap) => {
                        const invData = snap.val() || [];
                        setInventory(invData.filter(Boolean));
                    });

                    // Listener Khusus untuk Tenaga Ahli (dipisah dari pmc_data)
                    firebase.database().ref('pmc_experts').on('value', (snap) => {
                        let expData = snap.val() || [];
                        expData = expData.filter(Boolean);

                        // --- AUTO CLEANUP DUPLICATES ---
                        const uniqueMap = new Map();
                        let hasDuplicates = false;

                        expData.forEach(exp => {
                            if (!exp || !exp.name) return;
                            const key = exp.name.trim().toLowerCase().replace(/\s+/g, ' ');
                            if (!uniqueMap.has(key)) {
                                uniqueMap.set(key, { ...exp });
                            } else {
                                hasDuplicates = true;
                                const existing = uniqueMap.get(key);
                                // Merge data
                                if (!existing.phone && exp.phone) existing.phone = exp.phone;
                                if (!existing.bidangIlmu && exp.bidangIlmu) existing.bidangIlmu = exp.bidangIlmu;
                                if (!existing.jenjang && exp.jenjang) existing.jenjang = exp.jenjang;
                                if (!existing.perusahaan && exp.perusahaan) existing.perusahaan = exp.perusahaan;
                                
                                // Merge certificates
                                if (exp.certificates && Array.isArray(exp.certificates)) {
                                    if (!existing.certificates) existing.certificates = [];
                                    exp.certificates.forEach(c => {
                                        if (c && c.certName) {
                                            const cKey = c.certName.toLowerCase().replace(/\s+/g, ' ');
                                            const lKey = c.certLevel ? c.certLevel.toLowerCase().replace(/\s+/g, ' ') : '';
                                            
                                            const existingCertIndex = existing.certificates.findIndex(xc => {
                                                const xcKey = xc.certName ? xc.certName.toLowerCase().replace(/\s+/g, ' ') : '';
                                                const xlKey = xc.certLevel ? xc.certLevel.toLowerCase().replace(/\s+/g, ' ') : '';
                                                return xcKey === cKey && xlKey === lKey;
                                            });

                                            if (existingCertIndex !== -1) {
                                                // Update expired date if same name and same level
                                                if (c.expiredDate) {
                                                    existing.certificates[existingCertIndex].expiredDate = c.expiredDate;
                                                }
                                            } else {
                                                // Add as new certificate if name is new, or if name is same but level is different
                                                existing.certificates.push(c);
                                            }
                                        }
                                    });
                                }
                            }
                        });

                        const cleanedData = Array.from(uniqueMap.values());
                        if (hasDuplicates) {
                            // Sync cleaned data back to Firebase silently
                            firebase.database().ref('pmc_experts').set(cleanedData);
                        }
                        
                        setExperts(cleanedData);
                    });

                    // Listener Khusus untuk Penugasan Tenaga Ahli
                    firebase.database().ref('pmc_assignments').on('value', (snap) => {
                        const asgData = snap.val() || [];
                        setAssignments(asgData.filter(Boolean));
                    });

                    // Listener Khusus untuk LPSE List
                    firebase.database().ref('pmc_lpse_list').on('value', (snap) => {
                        const lpseData = snap.val();
                        if (lpseData && Array.isArray(lpseData)) {
                            setLpseList(lpseData);
                        } else {
                            // default list dari user request
                            const defaultLpse = [
                                "Kabupaten Badung", "Kabupaten Bangli", "Kabupaten Buleleng",
                                "Kabupaten Gianyar", "Kabupaten Jembrana", "Kabupaten Karangasem",
                                "Kabupaten Klungkung", "Kabupaten Tabanan", "Kota Denpasar",
                                "LPSE Kemen PU", "LPSE ATR", "LPSE Kemenpar", "LPSE Kemendikbud",
                                "LPSE Kemenkumham", "LPSE Kemendag", "LPSE Kemen LHK", "LPSE Kemenkeu",
                                "LPSE Polri", "LPSE BPOM", "LPSE Mah-Agung", "LPSE Kemen Perhubungan",
                                "LPSE INAPROC Nasional", "Kantor Desa Dauh Puri Kangin",
                                "Kantor Desa Dauh Puri Kelod", "Kantor Desa Saba Blahbatuh",
                                "Bank BPD", "Bank Mandiri", "LPSE Provinsi Bali"
                            ];
                            setLpseList(defaultLpse);
                            // Set ke firebase jika belum ada
                            firebase.database().ref('pmc_lpse_list').set(defaultLpse);
                        }
                    });

                    // Listener Khusus untuk Cert List
                    firebase.database().ref('pmc_cert_list').on('value', (snap) => {
                        const certData = snap.val();
                        if (certData && Array.isArray(certData)) {
                            setCertList(certData);
                        } else {
                            const defaultCerts = [
                                "Teknik Air Minum", "AMDAL", "Teknik Lingkungan", "STRA", "Arsitek", "Lanskap", "Teknik Bangunan Gedung", "Interior", "Teknik Elektronika dan Telekomunikasi dalam Gedung", "Geodesi", "Geodesi Jalan dan Jembatan", "Geodesi Bangunan Gedung", "Pengukuran Jalan", "Survei Terestris", "Geoteknik", "Penilai Bangunan Hijau", "Greenship", "Iluminasi", "Teknik Jalan", "Perencanaan Drainase", "Teknik Jembatan", "K3 Konstruksi", "Kesehatan Masyarakat", "Keselamatan Jalan", "Manajemen Konstruksi", "Manajemen Mutu", "Teknik Mekanikal", "Manajemen Proyek", "Perencanaan Wilayah dan Kota", "Teknik Plambing", "Teknik Proteksi Kebakaran", "Quantity Surveyor", "Teknik Sanitasi dan Limbah", "Sumber Daya Air", "Sistem Tata Udara", "Elektrikal Konstruksi Bangunan Gedung", "Teknik Transportasi dalam Gedung", "BIM", "Manajer Pengelolaan Bangunan Gedung", "Pengawas Bangunan Gedung", "Pelaksana Bangunan Gedung", "Pemeriksa Kelaikan Fungsi Struktur Bangunan Gedung", "Perawatan Bangunan Gedung", "Pelaksana Saluran Irigasi", "Teknik Bendungan Besar", "Sistem Informasi Geografis", "Kewilayahan", "Teknik Tenaga Listrik", "Teknik Sistem Tata Udara dan Refigrasi", "Pelaksana Pemeliharaan Jembatan"
                            ];
                            setCertList(defaultCerts);
                            firebase.database().ref('pmc_cert_list').set(defaultCerts);
                        }
                    });

                    // Pasang Listener Real-time

                    firebaseDbRef.on('value', (snap) => {
                        const data = snap.val() || { projects: [], resources: [], inventory: [] };

                        // DEKODE TRIK JSON UNTUK DETAIL TIM
                        const processedProjects = (data.projects || []).filter(Boolean).map(p => {
                            let parsedTeam = [];
                            let catDetails = {};
                            let leader = "";
                            let individualStatus = {};
                            let pengawasanDetails = {};
                            const rawString = Array.isArray(p.team) ? p.team.join(';') : (p.team || '');

                            if (rawString.startsWith('{')) {
                                try {
                                    const parsed = JSON.parse(rawString);
                                    parsedTeam = parsed.members || [];
                                    catDetails = parsed.details || {};
                                    leader = parsed.leader || "";
                                    individualStatus = parsed.individualStatus || {};
                                    pengawasanDetails = parsed.pengawasanDetails || {};
                                } catch (e) { parsedTeam = p.team || []; }
                            } else {
                                parsedTeam = p.team || [];
                                catDetails = p.categoryDetails || {};
                                leader = p.teamLeader || "";
                                individualStatus = p.individualStatus || {};
                                pengawasanDetails = p.pengawasanDetails || {};
                            }

                            const finalPengawasan = {};
                            for (const k in pengawasanDetails || {}) {
                                finalPengawasan[decodeKey(k)] = pengawasanDetails[k];
                            }
                            const finalIndividual = {};
                            for (const k in individualStatus || {}) {
                                finalIndividual[decodeKey(k)] = individualStatus[k];
                            }

                            return { ...p, team: parsedTeam, categoryDetails: catDetails, teamLeader: leader, individualStatus: finalIndividual, pengawasanDetails: finalPengawasan };
                        });

                        const processedResources = (data.resources || []).filter(Boolean).map(r => {
                            let role = r.role || "";
                            let level = "Staff";
                            let manualPoints = 0;
                            if (role.includes("|")) {
                                const parts = role.split("|");
                                role = parts[0];
                                level = parts[1] || "Staff";
                                manualPoints = parseInt(parts[2]) || 0;
                            }
                            return { ...r, role, level, manualPoints };
                        });

                        // Auto-Heal Project Names: Mencegah bug "Lainnya" akibat perbedaan tanda baca pada nama
                        const normalizeName = (name) => name ? name.toLowerCase().replace(/[^a-z0-9]/g, '') : '';
                        const healedProjects = processedProjects.map(p => {
                            const healedP = { ...p };
                            const healName = (oldName) => {
                                if (!oldName) return oldName;
                                const exactMatch = processedResources.find(r => r.name === oldName);
                                if (exactMatch) return oldName;
                                const fuzzyMatch = processedResources.find(r => normalizeName(r.name) === normalizeName(oldName));
                                return fuzzyMatch ? fuzzyMatch.name : oldName;
                            };

                            healedP.teamLeader = healName(healedP.teamLeader);
                            if (healedP.team) healedP.team = healedP.team.map(m => healName(m));

                            const newPengawasan = {};
                            for (const k in healedP.pengawasanDetails) newPengawasan[healName(k)] = healedP.pengawasanDetails[k];
                            healedP.pengawasanDetails = newPengawasan;

                            const newIndividual = {};
                            for (const k in healedP.individualStatus) newIndividual[healName(k)] = healedP.individualStatus[k];
                            healedP.individualStatus = newIndividual;

                            return healedP;
                        });

                        setProjects(healedProjects);
                        setResources(processedResources);
                        setIsLive(true);
                        setLoading(false);
                    }, (error) => {
                        console.error("Firebase Read Error:", error);
                        setErrorMsg("Koneksi Firebase terputus.");
                        setIsLive(false);
                        setLoading(false);
                    });
                }, (error) => {
                    console.error("Firebase Initial Read Error:", error);
                    let msg = "Koneksi Firebase gagal.";
                    if (error.code === 'PERMISSION_DENIED') {
                        msg = "Akses ditolak. Masa aktif test mode Firebase Anda sepertinya sudah habis. Harap perbarui rules database di konsol Firebase.";
                    }
                    setErrorMsg(msg);
                    setIsLive(false);
                    setLoading(false);
                });
            };


            const fetchFromGoogleSheets = () => {
                // Dipanggil oleh tombol refresh manual (Hanya memicu loading animasi sebentar)
                setLoading(true);
                setTimeout(() => setLoading(false), 500);
            };

            useEffect(() => { initFirebaseListener(); }, []);

            // Mengganti Tab dan me-reset view
            const handleTabChange = (tab) => {
                if ((tab === 'kpi' || tab === 'inventaris' || tab === 'ahli' || tab === 'penugasan') && !isDeveloper) {
                    setShowDevPrompt(true);
                    setPendingTabAttempt(tab);
                    setDevPassword("");
                    setDevAuthError("");
                    setSidebarOpen(false);
                    return;
                }
                setActiveTab(tab);
                if (tab !== 'schedule') {
                    setActiveScheduleProject(null);
                }
                setViewingEmployee(null);
                setSidebarOpen(false);

                // Reset semua filter pencarian saat pindah tab agar data tidak terlihat 'hilang'
                setSearchTeamTab("");
                setSearchProjectTab("");
                setSearchKPITab("");
                setSearchGanttTab("");
                setSearchInvTab("");
                setSearchExpertTab("");
                setSearchAssignmentTab("");
            };

            const handleDevLogin = () => {
                if (devPassword === "Batanghari15") {
                    setIsDeveloper(true);
                    setShowDevPrompt(false);
                    setActiveTab(pendingTabAttempt || 'kpi');
                    setPendingTabAttempt(null);
                } else {
                    setDevAuthError("Password salah! Akses ditolak.");
                }
            };

            const handleAnalyzeDomino = (project) => {
                const today = new Date();
                today.setHours(0, 0, 0, 0);

                let globalDelayDays = 0;
                if (project.deadline) {
                    const deadlineDate = new Date(project.deadline);
                    deadlineDate.setHours(0, 0, 0, 0);
                    globalDelayDays = Math.ceil((today.getTime() - deadlineDate.getTime()) / (1000 * 60 * 60 * 24));
                }

                const impactedEmployees = [];
                let maxDelay = globalDelayDays > 0 ? globalDelayDays : 0;

                if (project.team && Array.isArray(project.team)) {
                    project.team.forEach(memberName => {
                        const isIndividuallyDone = project.individualStatus?.[memberName] === true;
                        if (!isIndividuallyDone) {
                            const res = resources.find(r => r.name === memberName);
                            let memberDeadlineStr = null;

                            if (res) {
                                const effectiveCat = getEffectiveEmpCategory(project, res.name, res.role);
                                if (project.categoryDetails?.[effectiveCat]?.deadline) {
                                    memberDeadlineStr = project.categoryDetails[effectiveCat].deadline;
                                }
                            }
                            if (!memberDeadlineStr) memberDeadlineStr = project.deadline;

                            if (!memberDeadlineStr) return;

                            const memberDeadlineDate = new Date(memberDeadlineStr);
                            memberDeadlineDate.setHours(0, 0, 0, 0);

                            const delayDays = Math.ceil((today.getTime() - memberDeadlineDate.getTime()) / (1000 * 60 * 60 * 24));

                            // Jika secara individu tidak terlambat, lewati
                            if (delayDays <= 0) return;

                            if (delayDays > maxDelay) maxDelay = delayDays;

                            const futureProjects = [];
                            computedProjects.forEach(fp => {
                                if (fp.id === project.id) return;
                                if (fp.computedStatus === 'Done') return;
                                if (fp.team && fp.team.includes(memberName)) {
                                    const fpIndvDone = fp.individualStatus?.[memberName] === true;
                                    if (!fpIndvDone) {
                                        let fpDeadlineStr = null;
                                        if (res) {
                                            const fpEffectiveCat = getEffectiveEmpCategory(fp, res.name, res.role);
                                            if (fp.categoryDetails?.[fpEffectiveCat]?.deadline) {
                                                fpDeadlineStr = fp.categoryDetails[fpEffectiveCat].deadline;
                                            }
                                        }
                                        if (!fpDeadlineStr) fpDeadlineStr = fp.deadline;

                                        if (fpDeadlineStr) {
                                            const fDate = new Date(fpDeadlineStr);
                                            fDate.setHours(0, 0, 0, 0);
                                            if (!isNaN(fDate.getTime()) && fDate >= today) {
                                                const remaining = Math.ceil((fDate.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
                                                // Jika sisa waktu proyek masa depan lebih kecil atau sama dengan hari keterlambatan individu ini
                                                if (remaining <= delayDays + 3) { // buffer 3 hari
                                                    futureProjects.push({
                                                        name: fp.name,
                                                        deadlineStr: fpDeadlineStr,
                                                        remainingDays: remaining,
                                                        clashDays: delayDays - remaining
                                                    });
                                                }
                                            }
                                        }
                                    }
                                }
                            });

                            if (futureProjects.length > 0) {
                                impactedEmployees.push({
                                    name: memberName,
                                    futureProjects,
                                    delayDays: delayDays // Simpan keterlambatan spesifik individu
                                });
                            }
                        }
                    });
                }

                if (impactedEmployees.length === 0 && globalDelayDays <= 0) {
                    return; // Tidak ada delay sama sekali
                }

                setDominoAnalysis({
                    project,
                    delayDays: maxDelay,
                    impactedEmployees
                });
            };
            const handleTogglePendingSubmit = () => {
                if (!pendingProjectData) return;

                const now = new Date();
                const dateStr = formatDateTimeIndo(now.toISOString());

                const updatedProject = {
                    ...pendingProjectData,
                    isPending: true,
                    pendingReason: pendingReasonText,
                    pendingDate: dateStr
                };

                handleCrudAction('edit', 'project', updatedProject);
                setShowPendingModal(false);
                setPendingProjectData(null);
                setPendingReasonText("");
            };

            const handleResumeProject = (project) => {
                setResumeProjectData(project);
                setShowResumeModal(true);
            };

            const handleConfirmResumeProject = () => {
                if (!resumeProjectData) return;
                const updatedProject = {
                    ...resumeProjectData,
                    isPending: false,
                    pendingReason: "",
                    pendingDate: ""
                };
                handleCrudAction('edit', 'project', updatedProject);
                setShowResumeModal(false);
                setResumeProjectData(null);
            };

            // CRUD ACTION INVENTORY
            const handleInventoryAction = async (action, payload) => {
                setLoading(true);
                setErrorMsg("");
                try {
                    let newData = [...inventory];

                    if (action === 'add') {
                        newData.push(payload);
                    } else if (action === 'borrow-cart') {
                        newData = newData.map(item => {
                            if (payload.selectedIds.includes(item.id)) {
                                return { ...item, status: 'Menunggu Verifikasi', ...payload.data };
                            }
                            return item;
                        });
                        setBorrowCart([]);
                    } else if (action === 'edit' || action === 'borrow' || action === 'return' || action === 'extend') {
                        newData = newData.map(item => item.id === payload.id ? payload : item);
                    } else if (action === 'delete') {
                        newData = newData.filter(item => item.id !== payload.id);
                    }

                    await firebase.database().ref('pmc_inventory').set(newData);
                    setInventory(newData);
                    setModalConfig({ isOpen: false, type: null, mode: 'add', data: null });
                } catch (error) {
                    console.error("Firebase Inventory Write Error:", error);
                    setErrorMsg("Gagal menyimpan data inventaris.");
                }
                setLoading(false);
            };

            const handleImportExcel = (e) => {
                const file = e.target.files[0];
                if (!file) return;

                const reader = new FileReader();
                reader.onload = async (evt) => {
                    setLoading(true);
                    try {
                        const bstr = evt.target.result;
                        const wb = XLSX.read(bstr, { type: 'binary' });
                        const wsname = wb.SheetNames[0];
                        const ws = wb.Sheets[wsname];
                        const data = XLSX.utils.sheet_to_json(ws);

                        const expertsMap = new Map();

                        data.forEach(row => {
                            const name = row['Nama Personil'] || row['Nama'] || row['Nama Lengkap'] || row['Name'];
                            if (!name) return;

                            const nameStr = name.toString().trim();
                            const nameKey = nameStr.toLowerCase().replace(/\s+/g, ' ');

                            if (!expertsMap.has(nameKey)) {
                                const existingExpert = experts.find(e => {
                                    if (!e || !e.name) return false;
                                    return e.name.trim().toLowerCase().replace(/\s+/g, ' ') === nameKey;
                                });
                                
                                // Kategori is removed from template, default to existing or 'Internal'
                                const category = row['Kategori'] || row['Category'] || (existingExpert ? existingExpert.category : 'Internal');
                                const phone = row['Kontak'] || row['No HP'] || row['No. HP'] || row['Telepon'] || (existingExpert ? existingExpert.phone : '');
                                const status = row['Status'] || (existingExpert ? existingExpert.status : 'Tersedia');
                                
                                // Merge Jenjang Pendidikan and Bidang Ilmu
                                const jenjangBidang = row['Jenjang dan Bidang Ilmu'] || row['Bidang Ilmu'] || row['Jurusan'];
                                let parsedBidangIlmu = existingExpert ? existingExpert.bidangIlmu : '';
                                let parsedJenjang = existingExpert ? existingExpert.jenjang : '';
                                
                                if (jenjangBidang) {
                                    parsedBidangIlmu = jenjangBidang.toString().trim();
                                    // If old template was used with split columns, try to capture it. Otherwise leave jenjang empty.
                                    parsedJenjang = row['Jenjang Pendidikan'] ? row['Jenjang Pendidikan'].toString().trim() : '';
                                } else {
                                    parsedJenjang = row['Jenjang Pendidikan'] || (existingExpert ? existingExpert.jenjang : '');
                                }

                                const perusahaan = row['Perusahaan'] || row['Instansi'] || row['Asal Perusahaan'] || (existingExpert ? existingExpert.perusahaan : '');

                                expertsMap.set(nameKey, {
                                    id: existingExpert ? existingExpert.id : 'exp-' + Date.now().toString() + Math.random().toString(36).substr(2, 5),
                                    name: existingExpert ? existingExpert.name : nameStr,
                                    category: category.toString(),
                                    phone: phone.toString(),
                                    status: status.toString(),
                                    jenjang: parsedJenjang.toString(),
                                    bidangIlmu: parsedBidangIlmu.toString(),
                                    perusahaan: perusahaan.toString(),
                                    certificates: existingExpert ? [...(existingExpert.certificates || [])] : [],
                                    tenders: existingExpert ? [...(existingExpert.tenders || [])] : []
                                });
                            }

                            const currentExpert = expertsMap.get(nameKey);

                            // Parse Certificate Data
                            const certName = row['Sertifikat Keahlian'];
                            const certLevel = row['Jenjang'];
                            const rawDate = row['Masa Berlaku'];

                            let expiredDateStr = '';
                            if (rawDate) {
                                if (typeof rawDate === 'number') {
                                    // Excel serial date to JS Date (UTC)
                                    const utc_days = Math.floor(rawDate - 25569);
                                    const date_info = new Date(utc_days * 86400 * 1000);
                                    expiredDateStr = `${date_info.getUTCFullYear()}-${String(date_info.getUTCMonth() + 1).padStart(2, '0')}-${String(date_info.getUTCDate()).padStart(2, '0')}`;
                                } else {
                                    const parsed = new Date(rawDate);
                                    if (!isNaN(parsed.getTime())) {
                                        expiredDateStr = parsed.toISOString().split('T')[0];
                                    } else {
                                        // fallback for DD/MM/YYYY
                                        const parts = rawDate.toString().split(/[-/]/);
                                        if (parts.length === 3) {
                                            expiredDateStr = `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
                                        }
                                    }
                                }
                            }

                            if (certName) {
                                const certNameStr = certName.toString().trim();
                                const cKey = certNameStr.toLowerCase().replace(/\s+/g, ' ');
                                const newLevelStr = certLevel ? certLevel.toString().trim() : 'Ahli Muda';
                                const lKey = newLevelStr.toLowerCase().replace(/\s+/g, ' ');

                                const existingCertIndex = currentExpert.certificates.findIndex(c => {
                                    const xcKey = c.certName ? c.certName.toLowerCase().replace(/\s+/g, ' ') : '';
                                    const xlKey = c.certLevel ? c.certLevel.toLowerCase().replace(/\s+/g, ' ') : '';
                                    return xcKey === cKey && xlKey === lKey;
                                });

                                if (existingCertIndex !== -1) {
                                    // Update expired date if same name and same level
                                    if (expiredDateStr) {
                                        currentExpert.certificates[existingCertIndex].expiredDate = expiredDateStr;
                                    }
                                } else {
                                    // Add new certificate if name is new, or same name but different level
                                    currentExpert.certificates.push({
                                        certName: certNameStr,
                                        certLevel: newLevelStr,
                                        issuedDate: '',
                                        expiredDate: expiredDateStr
                                    });
                                }
                            }
                        });

                        const importedData = Array.from(expertsMap.values());

                        if (importedData.length > 0) {
                            if (confirm(`Ditemukan ${importedData.length} data tenaga ahli (baru/update) dari Excel. Lanjutkan sinkronisasi ke database?`)) {
                                const newExperts = [...experts];
                                
                                importedData.forEach(importedExp => {
                                    const index = newExperts.findIndex(e => e.id === importedExp.id);
                                    if (index !== -1) {
                                        newExperts[index] = importedExp;
                                    } else {
                                        newExperts.push(importedExp);
                                    }
                                });

                                await firebase.database().ref('pmc_experts').set(newExperts);
                                setExperts(newExperts);
                                alert('Sinkronisasi data tenaga ahli berhasil diselesaikan!');
                            }
                        } else {
                            alert('Tidak ada data valid yang ditemukan. Pastikan Excel Anda memiliki kolom "Nama Personil".');
                        }
                    } catch (err) {
                        console.error('Import error:', err);
                        alert('Gagal membaca atau memproses file Excel.');
                    } finally {
                        setLoading(false);
                        e.target.value = '';
                    }
                };
                reader.readAsBinaryString(file);
            };

            // CRUD ACTION EXPERTS
            const handleExpertAction = (action, payload) => {
                // setLoading(true);
                setErrorMsg("");
                try {
                    let newData = [...experts];

                    if (action === 'add') {
                        payload.id = 'exp-' + Date.now().toString();
                        if (!payload.certificates) payload.certificates = [];
                        if (!payload.tenders) payload.tenders = [];
                        newData.push(payload);
                    } else if (action === 'edit' || action === 'update_certificates' || action === 'update_tenders') {
                        newData = newData.map(item => item.id === payload.id ? payload : item);
                    } else if (action === 'delete') {
                        newData = newData.filter(item => item.id !== payload.id);
                    }

                    if (expertSaveTimeout) clearTimeout(expertSaveTimeout);
                    expertSaveTimeout = setTimeout(() => {
                        firebase.database().ref('pmc_experts').set(newData).catch(e => console.error(e));
                    }, 800);
                    setExperts(newData);
                    if (action !== 'delete') setModalConfig({ isOpen: false, type: null, mode: 'add', data: null });
                } catch (error) {
                    console.error("Firebase Expert Write Error:", error);
                    setErrorMsg("Gagal menyimpan data tenaga ahli.");
                }
                // setLoading(false);
            };


            // =========================================================================
            // SINKRONISASI OTOMATIS: PENUGASAN → LIST PROYEK (PENGAWASAN)
            // Fungsi ini otomatis membuat/memperbarui/menghapus proyek di List Proyek
            // saat user menyimpan data di menu Penugasan Tenaga Ahli.
            // Field `sourceAssignmentId` dipakai sebagai penghubung antar data.
            // =========================================================================
            const syncAssignmentToProject = async (assignment, action) => {
                if (!firebaseDbRef) return;
                try {
                    const snapshot = await firebaseDbRef.once('value');
                    const currentData = snapshot.val() || { projects: [], resources: [] };
                    let allProjects = (currentData.projects || []).filter(Boolean);

                    if (action === 'delete') {
                        // Hapus proyek yang terhubung
                        const linkedProject = allProjects.find(p => p.sourceAssignmentId === assignment.id);
                        if (linkedProject) {
                            allProjects = allProjects.filter(p => p.sourceAssignmentId !== assignment.id);
                            const cleanProjects = allProjects.map(proj => {
                                const newP = { ...proj };
                                if (newP.pengawasanDetails) {
                                    const cleanP = {};
                                    for (const k in newP.pengawasanDetails) cleanP[encodeKey(k)] = newP.pengawasanDetails[k];
                                    newP.pengawasanDetails = cleanP;
                                }
                                return newP;
                            });
                            await firebaseDbRef.update({ projects: cleanProjects });
                        }
                        return;
                    }

                    // Buat atau perbarui proyek dari data assignment
                    const expertNames = [...new Set((assignment.experts || []).map(exp => {
                        const expertObj = experts.find(e => e.id === exp.expertId);
                        return expertObj ? (expertObj.linkedResourceName || expertObj.name) : null;
                    }).filter(Boolean))];

                    // Bangun pengawasanDetails baru dari data experts di assignment
                    // Pertahankan statusTurun yang sudah ada sebelumnya
                    const existingProject = allProjects.find(p => p.sourceAssignmentId === assignment.id);
                    const existingPengawasanDetails = existingProject ? (() => {
                        // Decode existing details
                        const raw = existingProject.pengawasanDetails || {};
                        const decoded = {};
                        for (const k in raw) decoded[decodeKey(k)] = raw[k];
                        return decoded;
                    })() : {};

                    const newPengawasanDetails = {};
                    (assignment.experts || []).forEach(exp => {
                        const expertObj = experts.find(e => e.id === exp.expertId);
                        if (!expertObj) return;
                        const name = expertObj.linkedResourceName || expertObj.name;
                        const existingDetail = existingPengawasanDetails[name] || {};
                        // Pertahankan statusTurun yang sudah di-set user, default "Tidak Turun"
                        newPengawasanDetails[encodeKey(name)] = {
                            role: exp.role || 'Inspector',
                            manMonth: exp.manMonth || '',
                            statusTurun: existingDetail.statusTurun || 'Tidak Turun',
                            deadline: existingDetail.deadline || '',
                        };
                    });

                    // Serialisasi team data ke JSON string (sesuai format sistem)
                    const teamDataObj = {
                        members: expertNames,
                        details: {},
                        leader: '',
                        individualStatus: existingProject ? (existingProject.individualStatus || {}) : {},
                        pengawasanDetails: newPengawasanDetails
                    };

                    const projectPayload = {
                        name: assignment.jobName || 'Tanpa Nama',
                        client: assignment.clientName || assignment.lpseName || '',
                        type: assignment.projectType || 'Pengawasan',
                        status: 'On Progress',
                        spmk: assignment.startDate || '',
                        deadline: assignment.endDate || '',
                        sourceAssignmentId: assignment.id,
                        description: existingProject ? (existingProject.description || '') : '',
                        descriptionUpdatedAt: existingProject ? (existingProject.descriptionUpdatedAt || '') : '',
                        notStarted: existingProject ? (existingProject.notStarted || false) : false,
                        isPending: existingProject ? (existingProject.isPending || false) : false,
                        team: JSON.stringify(teamDataObj).replace(/;/g, ','),
                    };

                    let updatedProjects;
                    if (existingProject) {
                        // EDIT: perbarui proyek yang sudah ada
                        projectPayload.id = existingProject.id;
                        updatedProjects = allProjects.map(p => p.sourceAssignmentId === assignment.id ? projectPayload : p);
                    } else {
                        // ADD: buat proyek baru
                        projectPayload.id = 'sync-' + assignment.id;
                        updatedProjects = [...allProjects, projectPayload];
                    }

                    await firebaseDbRef.update({ projects: updatedProjects });
                } catch (err) {
                    console.error("Sync Assignment to Project Error:", err);
                }
            };

            // CRUD ACTION ASSIGNMENTS
            const handleAssignmentAction = (action, payload) => {
                setErrorMsg("");
                try {
                    let newData = [...assignments];

                    if (action === 'add') {
                        payload.id = 'asg-' + Date.now().toString();
                        newData.push(payload);
                    } else if (action === 'edit') {
                        newData = newData.map(item => item.id === payload.id ? payload : item);
                    } else if (action === 'delete') {
                        newData = newData.filter(item => item.id !== payload.id);
                    }

                    if (assignmentSaveTimeout) clearTimeout(assignmentSaveTimeout);
                    assignmentSaveTimeout = setTimeout(() => {
                        firebase.database().ref('pmc_assignments').set(newData).catch(error => {
                            console.error("Firebase Assignment Write Error:", error);
                        });
                    }, 800);

                    setAssignments(newData);
                    if (action !== 'delete') setModalConfig({ isOpen: false, type: null, mode: 'add', data: null });

                    // Sinkronisasi otomatis ke List Proyek (hanya untuk tipe Pengawasan)
                    const projectType = (payload.projectType || 'Pengawasan').toLowerCase();
                    if (projectType.includes('pengawas') || projectType.includes('manajemen konstruksi')) {
                        syncAssignmentToProject(payload, action);
                    }
                } catch (error) {
                    console.error("Local Assignment Update Error:", error);
                }
            };

            // CRUD ACTION FIREBASE REALTIME
            const handleCrudAction = async (action, type, payload) => {
                setLoading(true);
                setErrorMsg("");
                try {
                    // Tarik data terbaru dari Firebase sekali saja untuk dimodifikasi
                    const snapshot = await firebaseDbRef.once('value');
                    const currentData = snapshot.val() || { projects: [], resources: [] };

                    let newProjects = currentData.projects || [];
                    let newResources = currentData.resources || [];

                    // Bersihkan array dari null jika ada
                    newProjects = newProjects.filter(Boolean);
                    newResources = newResources.filter(Boolean);

                    if (action === 'add') {
                        payload.id = Date.now().toString();
                        if (type === 'project') newProjects.push(payload);
                        else newResources.push(payload);
                    } else if (action === 'edit') {
                        if (type === 'project') {
                            newProjects = newProjects.map(p => p.id === payload.id ? payload : p);
                        } else {
                            newResources = newResources.map(r => r.id === payload.id ? payload : r);
                        }
                    } else if (action === 'delete') {
                        if (type === 'project') {
                            newProjects = newProjects.filter(p => p.id !== payload.id);
                        } else {
                            newResources = newResources.filter(r => r.id !== payload.id);
                        }
                    }

                    const cleanProjects = newProjects.map(proj => {
                        const newP = { ...proj };
                        if (newP.pengawasanDetails) {
                            const cleanP = {};
                            for (const k in newP.pengawasanDetails) {
                                cleanP[encodeKey(k)] = newP.pengawasanDetails[k];
                            }
                            newP.pengawasanDetails = cleanP;
                        }
                        if (newP.individualStatus) {
                            const cleanI = {};
                            for (const k in newP.individualStatus) {
                                cleanI[encodeKey(k)] = newP.individualStatus[k];
                            }
                            newP.individualStatus = cleanI;
                        }
                        return newP;
                    });

                    await firebaseDbRef.update({
                        projects: cleanProjects,
                        resources: newResources
                    });

                    // Auto create expert if adding a new resource
                    if (action === 'add' && type !== 'project') {
                        try {
                            const expRef = firebase.database().ref('pmc_experts');
                            const expSnap = await expRef.once('value');
                            const expList = (expSnap.val() || []).filter(Boolean);
                            
                            // Check if expert with exact name already exists
                            if (!expList.some(e => e.name === payload.name)) {
                                expList.push({
                                    id: 'exp-' + Date.now().toString(),
                                    name: payload.name,
                                    linkedResourceName: payload.name,
                                    phone: '',
                                    status: 'Tersedia',
                                    jenjang: '',
                                    bidangIlmu: '',
                                    perusahaan: '',
                                    certificates: [],
                                    tenders: []
                                });
                                await expRef.set(expList);
                            }
                        } catch (expError) {
                            console.error("Failed to auto-create expert from resource:", expError);
                        }
                    }

                    closeModal();
                    // Tidak perlu memanggil fetch ulang, karena on('value') akan otomatis merender state seketika!
                } catch (error) {
                    console.error("Error Action Firebase:", error);
                    setErrorMsg(error.message || "Terjadi kesalahan saat menyimpan data ke Firebase.");
                } finally {
                    setLoading(false);
                }
            };

            const handleToggleIndividualStatus = (projectId, employeeName, newStatus) => {
                setConfirmDialog({
                    isOpen: true,
                    type: newStatus ? 'info' : 'danger',
                    title: newStatus ? 'Konfirmasi Selesai' : 'Batalkan Status Selesai',
                    message: newStatus
                        ? `Tandai tugas khusus untuk ${employeeName} sebagai Selesai di proyek ini?`
                        : `Batalkan status Selesai untuk ${employeeName}?`,
                    onConfirm: async () => {
                        const project = projects.find(p => p.id === projectId);
                        if (!project) return;

                        const newIndividualStatus = { ...(project.individualStatus || {}) };
                        if (newStatus) {
                            newIndividualStatus[employeeName] = true;
                        } else {
                            delete newIndividualStatus[employeeName];
                        }

                        const teamDataObj = {
                            members: project.team || [],
                            details: project.categoryDetails || {},
                            leader: project.teamLeader || "",
                            individualStatus: newIndividualStatus
                        };

                        const updatedPayload = { ...project, team: JSON.stringify(teamDataObj) };

                        delete updatedPayload.categoryDetails;
                        delete updatedPayload.individualStatus;
                        delete updatedPayload.teamLeader;
                        delete updatedPayload.computedStatus;

                        await handleCrudAction('edit', 'project', updatedPayload);
                    }
                });
            };

            const handleToggleNotStarted = (projectId, newNotStartedStatus) => {
                setConfirmDialog({
                    isOpen: true,
                    type: newNotStartedStatus ? 'warning' : 'info',
                    title: newNotStartedStatus ? 'Tandai Proyek Belum Mulai' : 'Mulai Proyek',
                    message: newNotStartedStatus
                        ? 'Tandai proyek ini sebagai "Belum Mulai"? Proyek ini tidak akan dihitung dalam beban kerja dan tidak muncul dalam Rincian Penugasan Pegawai.'
                        : 'Tandai proyek ini sebagai "Sudah Mulai"? Proyek ini akan kembali dihitung dalam beban kerja dan muncul dalam KPI personil.',
                    onConfirm: async () => {
                        const project = projects.find(p => p.id === projectId);
                        if (!project) return;

                        const teamDataObj = {
                            members: project.team || [],
                            details: project.categoryDetails || {},
                            leader: project.teamLeader || "",
                            individualStatus: project.individualStatus || {},
                            pengawasanDetails: project.pengawasanDetails || {}
                        };

                        const updatedPayload = { ...project, team: JSON.stringify(teamDataObj), notStarted: newNotStartedStatus };

                        delete updatedPayload.categoryDetails;
                        delete updatedPayload.individualStatus;
                        delete updatedPayload.teamLeader;
                        delete updatedPayload.computedStatus;
                        delete updatedPayload.pengawasanDetails;

                        await handleCrudAction('edit', 'project', updatedPayload);
                    }
                });
            };

            const openModal = (type, mode, data = null) => {
                let parsedData = data;
                if (type === 'project' && data) {
                    let normalizedStatus = data.status === "On Track" ? "On Progress" : data.status;

                    if (data.deadline) {
                        try {
                            const d = new Date(data.deadline);
                            if (!isNaN(d.getTime())) {
                                const yyyy = d.getFullYear();
                                const mm = String(d.getMonth() + 1).padStart(2, '0');
                                const dd = String(d.getDate()).padStart(2, '0');
                                parsedData = { ...data, deadline: `${yyyy}-${mm}-${dd}`, status: normalizedStatus };
                            } else {
                                parsedData = { ...data, status: normalizedStatus };
                            }
                        } catch (e) { parsedData = { ...data, status: normalizedStatus }; }
                    } else {
                        parsedData = { ...data, status: normalizedStatus };
                    }
                }
                setModalConfig({ isOpen: true, type, mode, data: parsedData });
            };

            const closeModal = () => setModalConfig({ isOpen: false, type: null, mode: 'add', data: null });

            const handleDelete = (type, id) => {
                setConfirmDialog({
                    isOpen: true,
                    type: 'danger',
                    title: 'Hapus Data Permanen',
                    message: 'Apakah Anda yakin ingin menghapus data ini secara permanen? Tindakan ini tidak dapat dibatalkan.',
                    onConfirm: () => {
                        handleCrudAction('delete', type, { id });
                    }
                });
            };

            const ErrorBanner = () => {
                if (!errorMsg) return null;
                return (
                    <div className="bg-red-50 border border-red-200 text-red-700 p-4 rounded-xl flex items-start gap-3 mb-6 shadow-sm fade-in">
                        <Icon name="alert-triangle" className="shrink-0 mt-0.5 text-red-500" />
                        <div>
                            <h4 className="font-bold text-sm">Masalah Koneksi Database</h4>
                            <p className="text-sm mt-1">{errorMsg}</p>
                            <p className="text-xs mt-2 text-red-600 font-medium">Solusi: Pastikan Anda telah menimpa seluruh kode backend (Code.gs) dengan yang terbaru dan melakukan Deploy "Versi Baru".</p>
                        </div>
                    </div>
                );
            };

            // --- KOMPONEN TAB: KPI & EVALUASI ---
            const renderKPI = () => {
                const leaders = resources.filter(r => r.level === 'Team Leader');
                const kordinators = resources.filter(r => r.level?.startsWith('Kordinator Divisi'));
                const staffs = resources.filter(r => r.level !== 'Team Leader' && !r.level?.startsWith('Kordinator Divisi'));

                const allKpiDataLeaders = leaders.map(res => {
                    const kpi = calculateLeaderKPI(res, computedProjects);
                    return { ...res, ...kpi };
                });

                const allKpiDataKordinators = kordinators.map(res => {
                    const kpi = calculateLeaderKPI(res, computedProjects);
                    return { ...res, ...kpi };
                });

                const allKpiDataStaffs = staffs.map(res => {
                    const kpi = calculateEmployeeKPI(res, computedProjects);
                    return { ...res, ...kpi };
                });

                const sortKpiData = (a, b) => {
                    // 1. Prioritas Utama: Skor KPI Tertinggi
                    if (b.score !== a.score) return b.score - a.score;

                    // 2. Prioritas Kedua (Tie-breaker 1): Rating Bintang Manual
                    if ((b.rating || 0) !== (a.rating || 0)) return (b.rating || 0) - (a.rating || 0);

                    // 3. Prioritas Ketiga (Tie-breaker 2): Jumlah Proyek Selesai Tepat Waktu Terbanyak
                    const doneA = (a.bonusDoneLeader !== undefined ? a.bonusDoneLeader : (a.bonusDone || 0)) + (a.bonusDoneStaff || 0);
                    const doneB = (b.bonusDoneLeader !== undefined ? b.bonusDoneLeader : (b.bonusDone || 0)) + (b.bonusDoneStaff || 0);
                    if (doneB !== doneA) return doneB - doneA;

                    // 4. Prioritas Keempat (Tie-breaker 3): Keterlambatan Paling Sedikit
                    if (a.delayed !== b.delayed) return a.delayed - b.delayed;

                    // 5. Prioritas Kelima (Tie-breaker 4): Rata-rata Progress Tertinggi
                    return b.avgProgress - a.avgProgress;
                };

                const kpiDataLeaders = allKpiDataLeaders
                    .filter(res => res.name.toLowerCase().includes(searchKPITab.toLowerCase()) || res.role.toLowerCase().includes(searchKPITab.toLowerCase()))
                    .sort(sortKpiData);

                const kpiDataKordinators = allKpiDataKordinators
                    .filter(res => res.name.toLowerCase().includes(searchKPITab.toLowerCase()) || res.role.toLowerCase().includes(searchKPITab.toLowerCase()))
                    .sort(sortKpiData);

                const kpiDataStaffs = allKpiDataStaffs
                    .filter(res => res.name.toLowerCase().includes(searchKPITab.toLowerCase()) || res.role.toLowerCase().includes(searchKPITab.toLowerCase()))
                    .sort(sortKpiData);

                const allKpiCombined = [...allKpiDataLeaders, ...allKpiDataKordinators, ...allKpiDataStaffs];
                const needsAttention = allKpiCombined.filter(k => k.score < 60 || (k.isOverloaded && k.delayed > 0));
                const healthyCount = allKpiCombined.length - needsAttention.length;

                const renderTableRows = (data, isLeaderTable) => {
                    if (data.length === 0) {
                        return <tr><td colSpan="5" className="p-8 text-center text-slate-400">Belum ada data personil.</td></tr>;
                    }
                    return data.map((kpi, idx) => (
                        <tr key={kpi.id} className="hover:bg-indigo-50/50 dark:hover:bg-indigo-900/20 transition-colors">
                            <td className="p-4 flex items-center gap-3">
                                <div className={`w-8 h-8 rounded-full flex items-center justify-center font-bold text-white shrink-0 ${idx === 0 ? 'bg-amber-400 shadow-lg shadow-amber-400/30' : idx === 1 ? 'bg-slate-300' : idx === 2 ? 'bg-amber-600' : 'bg-slate-800'}`}>
                                    {idx + 1}
                                </div>
                                <div>
                                    <div className="flex items-center gap-2">
                                        <p className="font-bold text-slate-800 dark:text-slate-200">{kpi.name}</p>
                                        {isLeaderTable && <span className="text-[9px] font-bold text-amber-600 bg-amber-100 px-1.5 py-0.5 rounded flex items-center gap-1"><Icon name="star" size={10} className="fill-amber-500" /> Leader</span>}
                                    </div>
                                    {kpi.level && kpi.level.startsWith('Kordinator Divisi') ? (
                                        <div className="text-xs text-slate-500">
                                            <p>{kpi.level}</p>
                                            <p>{kpi.role}</p>
                                        </div>
                                    ) : (
                                        <p className="text-xs text-slate-500">
                                            {kpi.level === 'PIC' ? `PIC ${kpi.role}` : kpi.role}
                                        </p>
                                    )}
                                </div>
                            </td>
                            <td className="p-4 text-center">
                                <span className={`inline-block px-3 py-1 rounded-full text-xs font-bold border ${kpi.score >= 80 ? 'bg-emerald-50 text-emerald-700 border-emerald-200' : kpi.score >= 60 ? 'bg-blue-50 text-blue-700 border-blue-200' : 'bg-red-50 text-red-700 border-red-200'}`}>
                                    {kpi.score}
                                </span>
                            </td>
                            <td className="p-4">
                                <div className="flex justify-center items-center gap-0.5 text-amber-400">
                                    {[...Array(5)].map((_, i) => (
                                        <svg key={i} xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill={i < kpi.rating ? "currentColor" : "none"} stroke={i < kpi.rating ? "currentColor" : "#cbd5e1"} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2" /></svg>
                                    ))}
                                </div>
                            </td>
                            <td className="p-4 text-center">
                                <span className="font-semibold text-slate-700 dark:text-slate-300">{kpi.avgProgress}%</span>
                            </td>
                            <td className="p-4">
                                <div className="flex flex-col gap-1 items-start">
                                    {isLeaderTable && kpi.bonusDoneLeader > 0 && <span className="text-[10px] bg-emerald-50 text-emerald-700 border border-emerald-200 px-2 py-0.5 rounded font-medium inline-block whitespace-nowrap">+{kpi.bonusDoneLeader * 20} Bonus Leader (Tepat Waktu)</span>}
                                    {isLeaderTable && kpi.bonusDoneStaff > 0 && <span className="text-[10px] bg-emerald-50 text-emerald-700 border border-emerald-200 px-2 py-0.5 rounded font-medium inline-block whitespace-nowrap">+{kpi.bonusDoneStaff * 15} Bonus Staff (Tepat Waktu)</span>}
                                    {!isLeaderTable && kpi.bonusDone > 0 && <span className="text-[10px] bg-emerald-50 text-emerald-700 border border-emerald-200 px-2 py-0.5 rounded font-medium inline-block whitespace-nowrap">+{kpi.bonusDone * 15} Bonus Selesai Tepat Waktu</span>}
                                    {kpi.isOverloaded && kpi.delayed > 0 && <span className="text-[10px] bg-red-50 text-red-600 border border-red-200 px-2 py-0.5 rounded font-medium inline-block whitespace-nowrap">Overload & Terlambat</span>}
                                    {kpi.isOverloaded && kpi.delayed === 0 && <span className="text-[10px] bg-purple-50 text-purple-700 border border-purple-200 px-2 py-0.5 rounded font-medium inline-block whitespace-nowrap">Performa Ekstra (&gt;4 Proyek)</span>}
                                    {kpi.delayed > 0 && <span className="text-[10px] bg-amber-50 text-amber-700 border border-amber-200 px-2 py-0.5 rounded font-medium inline-block whitespace-nowrap">{kpi.delayed} Proyek Terlambat</span>}
                                    {kpi.atRisk > 0 && <span className="text-[10px] bg-slate-100 text-slate-600 border border-slate-200 px-2 py-0.5 rounded font-medium inline-block whitespace-nowrap">{kpi.atRisk} Proyek Beresiko</span>}
                                    {!kpi.isOverloaded && kpi.delayed === 0 && kpi.atRisk === 0 && <span className="text-xs text-slate-400 italic">Bersih & Aman</span>}
                                </div>
                            </td>
                        </tr>
                    ));
                };

                return (
                    <div className="space-y-6 fade-in">
                        <div className="bg-indigo-50 border border-indigo-200 text-indigo-800 p-4 rounded-xl flex items-start gap-3 shadow-sm">
                            <Icon name="alert-triangle" className="shrink-0 mt-0.5 text-indigo-600" />
                            <div>
                                <h4 className="font-bold text-sm">Tahap Pengembangan (BETA)</h4>
                                <p className="text-sm mt-1">Menu Evaluasi dan KPI ini belum resmi digunakan dan masih dalam tahap uji coba (pengembangan logika). Data di bawah ini belum menjadi hasil penilaian akhir.</p>
                            </div>
                        </div>

                        <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 glass-card p-6 rounded-2xl">
                            <div>
                                <h3 className="text-xl font-bold text-slate-800 dark:text-slate-100">Evaluasi & KPI Personil (BETA)</h3>
                                <p className="text-sm text-slate-500 mt-1">Mengukur kinerja berdasarkan ketepatan waktu, manajemen beban, dan rating.</p>
                            </div>
                            <div className="flex gap-2">
                                <div className="px-4 py-2 bg-emerald-50 text-emerald-700 rounded-lg border border-emerald-100 flex items-center gap-2 shrink-0">
                                    <Icon name="check-circle-2" size={18} />
                                    <span className="font-bold text-sm">{healthyCount} Sehat</span>
                                </div>
                                <div className="px-4 py-2 bg-amber-50 text-amber-700 rounded-lg border border-slate-200 flex items-center gap-2 shrink-0">
                                    <Icon name="alert-triangle" size={18} />
                                    <span className="font-bold text-sm">{needsAttention.length} Perlu Pantauan</span>
                                </div>
                            </div>
                        </div>

                        <div className="flex flex-row gap-3 items-center bg-white/70 dark:bg-slate-800/60 backdrop-blur-2xl p-4 rounded-2xl shadow-xl shadow-slate-200/50 dark:shadow-slate-900/50 border border-white/60 dark:border-slate-700/50">
                            <div className="relative flex-1 sm:flex-none sm:w-96">
                                <Icon name="search" size={18} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                                <input
                                    type="text"
                                    placeholder="Cari nama atau Tim..."
                                    value={searchKPITab}
                                    onChange={(e) => setSearchKPITab(e.target.value)}
                                    className="w-full pl-10 pr-4 py-2 bg-slate-50 dark:bg-slate-900/50 border border-slate-200 dark:border-slate-700 rounded-xl text-sm focus:outline-none focus:border-blue-500 focus:bg-white dark:focus:bg-slate-900 transition-colors dark:text-slate-200"
                                />
                            </div>
                            <button
                                onClick={() => setShowKPIInfoModal(true)}
                                className="shrink-0 p-2.5 text-slate-400 hover:text-blue-500 bg-slate-50 hover:bg-blue-50 dark:bg-slate-900/50 dark:hover:bg-blue-900/30 rounded-xl transition-all border border-slate-200 dark:border-slate-700 shadow-sm hover:shadow-md"
                                title="Informasi Perhitungan KPI"
                            >
                                <Icon name="help-circle" size={20} />
                            </button>
                        </div>

                        {/* TABEL TEAM LEADER */}
                        <div className="bg-white/70 dark:bg-slate-800/60 backdrop-blur-2xl rounded-2xl border border-white/60 dark:border-amber-900/30 shadow-xl shadow-slate-200/50 dark:shadow-slate-900/50 overflow-hidden flex flex-col mt-6">
                            <div className="p-4 bg-amber-50 dark:bg-amber-900/20 border-b border-amber-200 dark:border-amber-900/30 flex items-center gap-2">
                                <Icon name="star" size={20} className="fill-amber-500 text-amber-600" />
                                <h3 className="font-bold text-amber-900 dark:text-amber-400">Leaderboard: Team Leader</h3>
                            </div>
                            <div className="p-0 overflow-x-auto min-h-[150px]">
                                <table className="enterprise-table text-left min-w-[700px]">
                                    <thead>
                                        <tr className="bg-slate-50 dark:bg-slate-900/50 text-slate-500 dark:text-slate-400 text-xs uppercase tracking-wider">
                                            <th className="p-4 font-semibold whitespace-nowrap">Peringkat & Personil</th>
                                            <th className="p-4 font-semibold text-center whitespace-nowrap">Skor KPI</th>
                                            <th className="p-4 font-semibold text-center whitespace-nowrap">Rating Manual</th>
                                            <th className="p-4 font-semibold text-center whitespace-nowrap">Progress Makro Rata-rata</th>
                                            <th className="p-4 font-semibold whitespace-nowrap">Rincian Masalah</th>
                                        </tr>
                                    </thead>
                                    <tbody className="divide-y divide-slate-100 dark:divide-slate-700/50 text-sm [&>tr]:transition-colors [&>tr:hover]:bg-slate-50/50 dark:[&>tr:hover]:bg-slate-800/30">
                                        {renderTableRows(kpiDataLeaders, true)}
                                    </tbody>
                                </table>
                            </div>
                        </div>

                        {/* TABEL KORDINATOR DIVISI */}
                        <div className="bg-white/70 dark:bg-slate-800/60 backdrop-blur-2xl rounded-2xl border border-white/60 dark:border-blue-900/30 shadow-xl shadow-slate-200/50 dark:shadow-slate-900/50 overflow-hidden flex flex-col mt-6">
                            <div className="p-4 bg-blue-50 dark:bg-blue-900/20 border-b border-blue-200 dark:border-blue-900/30 flex items-center gap-2">
                                <Icon name="target" size={20} className="text-blue-600 dark:text-blue-400" />
                                <h3 className="font-bold text-blue-900 dark:text-blue-400">Leaderboard: Kordinator Divisi</h3>
                            </div>
                            <div className="p-0 overflow-x-auto min-h-[150px]">
                                <table className="enterprise-table text-left min-w-[700px]">
                                    <thead>
                                        <tr className="bg-slate-50 dark:bg-slate-900/50 text-slate-500 dark:text-slate-400 text-xs uppercase tracking-wider">
                                            <th className="p-4 font-semibold whitespace-nowrap">Peringkat & Personil</th>
                                            <th className="p-4 font-semibold text-center whitespace-nowrap">Skor KPI</th>
                                            <th className="p-4 font-semibold text-center whitespace-nowrap">Rating Manual</th>
                                            <th className="p-4 font-semibold text-center whitespace-nowrap">Progress Makro Rata-rata</th>
                                            <th className="p-4 font-semibold whitespace-nowrap">Rincian Masalah</th>
                                        </tr>
                                    </thead>
                                    <tbody className="divide-y divide-slate-100 dark:divide-slate-700/50 text-sm [&>tr]:transition-colors [&>tr:hover]:bg-slate-50/50 dark:[&>tr:hover]:bg-slate-800/30">
                                        {renderTableRows(kpiDataKordinators, true)}
                                    </tbody>
                                </table>
                            </div>
                        </div>

                        {/* TABEL STAFF */}
                        <div className="glass-card rounded-2xl overflow-hidden flex flex-col mt-6">
                            <div className="p-4 bg-slate-50 dark:bg-slate-900/50 border-b border-slate-200 dark:border-slate-700/50 flex items-center gap-2">
                                <Icon name="users" size={20} className="text-slate-600 dark:text-slate-400" />
                                <h3 className="font-bold text-slate-800 dark:text-slate-100">Leaderboard: Staff & Anggota Tim</h3>
                            </div>
                            <div className="p-0 overflow-x-auto min-h-[300px]">
                                <table className="enterprise-table text-left min-w-[700px]">
                                    <thead>
                                        <tr className="bg-slate-50 dark:bg-slate-900/50 text-slate-500 dark:text-slate-400 text-xs uppercase tracking-wider">
                                            <th className="p-4 font-semibold whitespace-nowrap">Peringkat & Personil</th>
                                            <th className="p-4 font-semibold text-center whitespace-nowrap">Skor KPI</th>
                                            <th className="p-4 font-semibold text-center whitespace-nowrap">Rating Manual</th>
                                            <th className="p-4 font-semibold text-center whitespace-nowrap">Progress Sub-Tim Rata-rata</th>
                                            <th className="p-4 font-semibold whitespace-nowrap">Rincian Masalah</th>
                                        </tr>
                                    </thead>
                                    <tbody className="divide-y divide-slate-100 dark:divide-slate-700/50 text-sm [&>tr]:transition-colors [&>tr:hover]:bg-slate-50/50 dark:[&>tr:hover]:bg-slate-800/30">
                                        {renderTableRows(kpiDataStaffs, false)}
                                    </tbody>
                                </table>
                            </div>
                        </div>

                    </div>
                );
            };

            // --- KOMPONEN TAB: DASHBOARD ---
            const renderDashboard = () => {
const statusPriority = { "Terlambat": 1, "Beresiko": 2, "On Progress": 3, "Done": 4 };
                const sortedProjectsForDashboard = [...computedProjects].sort((a, b) =>
                    statusPriority[a.computedStatus] - statusPriority[b.computedStatus]
                );
                // Menghitung Distribusi Sub-Tim
                const subTeamCounts = {};
                resources.forEach(r => {
                    const cat = getCategoryFromRole(r.role);
                    subTeamCounts[cat] = (subTeamCounts[cat] || 0) + 1;
                });
                
                // Urutkan dan ambil top 4, sisanya masuk 'Lainnya'
                const sortedSubTeams = Object.entries(subTeamCounts).sort((a, b) => b[1] - a[1]);
                const topSubTeams = sortedSubTeams.slice(0, 4);
                const otherSubTeamsCount = sortedSubTeams.slice(4).reduce((sum, [_, count]) => sum + count, 0);
                
                if (otherSubTeamsCount > 0) {
                    topSubTeams.push(['Lainnya', otherSubTeamsCount]);
                }

                const totalSubTeamMembers = resources.length || 1;
                const badgeColors = ['bg-indigo-500', 'bg-blue-500', 'bg-emerald-500', 'bg-amber-500', 'bg-slate-400'];
                const hexColors = ['#6366f1', '#3b82f6', '#10b981', '#f59e0b', '#94a3b8'];
                
                let currentAngle = 0;
                const conicStops = topSubTeams.map(([_name, count], i) => {
                    const angle = (count / totalSubTeamMembers) * 100;
                    const start = currentAngle;
                    const end = currentAngle + angle;
                    currentAngle = end;
                    return `${hexColors[i]} ${start}% ${end}%`;
                }).join(', ');

                return (
                    <div className="space-y-8 fade-in">
                        <ErrorBanner />
                        
                        <div className="grid grid-cols-1 md:grid-cols-4 lg:grid-cols-6 gap-4 lg:gap-6">
                            {/* BENTO CELL 1: Total Proyek (Hero) */}
                            <div 
                                onClick={() => setShowProjectTypeModal(true)}
                                className="col-span-1 md:col-span-2 lg:col-span-2 lg:row-span-2 bg-gradient-to-br from-indigo-600 to-blue-700 dark:from-indigo-900 dark:to-blue-900 rounded-3xl p-6 lg:p-8 shadow-xl shadow-indigo-200/50 dark:shadow-none text-white relative overflow-hidden cursor-pointer hover:scale-[1.02] transition-transform duration-300 flex flex-col justify-between group"
                            >
                                <div className="absolute top-0 right-0 p-6 opacity-20 group-hover:scale-110 transition-transform duration-500">
                                    <Icon name="briefcase" size={120} />
                                </div>
                                <div className="relative z-10">
                                    <p className="text-indigo-100 font-medium mb-1 lg:mb-2 text-sm lg:text-base">Total Keseluruhan</p>
                                    <h3 className="text-5xl lg:text-7xl font-black">{computedProjects.length}</h3>
                                    <p className="text-lg lg:text-2xl font-bold mt-2">Proyek Terdaftar</p>
                                </div>
                                <div className="relative z-10 mt-8 inline-flex items-center gap-2 text-sm font-medium bg-white/20 hover:bg-white/30 transition-colors px-4 py-2 rounded-full backdrop-blur-md w-max">
                                    <Icon name="folder-open" size={16} /> Lihat Detail Tipe
                                </div>
                            </div>

                            {/* BENTO CELL 2: Proyek Aktif & Selesai */}
                            <div className="col-span-1 md:col-span-2 lg:col-span-2 bg-white/80 dark:bg-slate-800/80 backdrop-blur-3xl rounded-[2rem] border border-slate-100 dark:border-slate-700/50 p-5 shadow-[0_8px_30px_rgb(0,0,0,0.04)] dark:shadow-[0_8px_30px_rgba(0,0,0,0.1)] flex gap-4 divide-x divide-slate-100 dark:divide-slate-700 transition-colors">
                                <div className="flex-1 flex flex-col justify-center">
                                    <div className="flex items-center gap-2 mb-3">
                                        <div className="w-8 h-8 rounded-full bg-indigo-50 dark:bg-indigo-900/50 flex items-center justify-center text-indigo-600 dark:text-indigo-400">
                                            <Icon name="target" size={16} />
                                        </div>
                                        <span className="text-sm font-semibold text-slate-600 dark:text-slate-300">Aktif</span>
                                    </div>
                                    <p className="text-3xl font-bold text-slate-800 dark:text-slate-100">{computedProjects.length - completedProjectsCount}</p>
                                    <p className="text-xs text-slate-500 mt-1">Sedang berjalan</p>
                                </div>
                                <div className="flex-1 flex flex-col justify-center pl-4">
                                    <div className="flex items-center gap-2 mb-3">
                                        <div className="w-8 h-8 rounded-full bg-emerald-50 dark:bg-emerald-900/50 flex items-center justify-center text-emerald-600 dark:text-emerald-400">
                                            <Icon name="check-circle-2" size={16} />
                                        </div>
                                        <span className="text-sm font-semibold text-slate-600 dark:text-slate-300">Selesai</span>
                                    </div>
                                    <p className="text-3xl font-bold text-slate-800 dark:text-slate-100">{completedProjectsCount}</p>
                                    <p className="text-xs text-slate-500 mt-1">Status Done</p>
                                </div>
                            </div>
                            {/* BENTO CELL 3: Distribusi Sub-Tim Pie Chart */}
                            <div 
                                onClick={() => setActiveTab('tim')}
                                className="col-span-1 md:col-span-2 lg:col-span-2 lg:row-span-2 glass-card rounded-[2rem] p-6 flex flex-col justify-between cursor-pointer hover:scale-[1.02] transition-transform duration-300 relative overflow-hidden"
                            >
                                <div className="flex justify-between items-center mb-4 relative z-10">
                                    <div>
                                        <h3 className="text-base font-bold text-slate-800 dark:text-slate-100">Distribusi Sub-Tim</h3>
                                        <p className="text-xs text-slate-500 mt-1">{resources.length} Personil Aktif</p>
                                    </div>
                                    <div className="w-10 h-10 rounded-full bg-blue-50 dark:bg-blue-900/30 flex items-center justify-center text-blue-600 dark:text-blue-400"><Icon name="pie-chart" size={20} /></div>
                                </div>
                                <div className="flex flex-col items-center flex-1 justify-center relative z-10 gap-4 mt-2">
                                    {resources.length > 0 ? (
                                        <div className="flex items-center gap-6 w-full justify-center">
                                            <div 
                                                className="w-24 h-24 sm:w-28 sm:h-28 rounded-full shrink-0 shadow-inner relative flex items-center justify-center"
                                                style={{ background: `conic-gradient(${conicStops})` }}
                                            >
                                                {/* Hole for donut chart effect */}
                                                <div className="w-14 h-14 sm:w-16 sm:h-16 bg-white dark:bg-slate-800 rounded-full flex flex-col items-center justify-center shadow-sm">
                                                    <span className="text-[10px] font-bold text-slate-400">Total</span>
                                                    <span className="text-sm font-black text-slate-700 dark:text-slate-200">{resources.length}</span>
                                                </div>
                                            </div>
                                            <div className="flex flex-col gap-2 max-w-[120px]">
                                                {topSubTeams.map(([name, count], i) => (
                                                    <div key={name} className="flex items-center gap-2">
                                                        <div className={`w-2.5 h-2.5 rounded-full shrink-0 ${badgeColors[i]}`}></div>
                                                        <span className="text-[10px] font-bold text-slate-600 dark:text-slate-300 truncate" title={name}>{name}</span>
                                                        <span className="text-[10px] text-slate-500 ml-auto">{count}</span>
                                                    </div>
                                                ))}
                                            </div>
                                        </div>
                                    ) : (
                                        <p className="text-xs text-slate-400 italic">Belum ada personil</p>
                                    )}
                                </div>
                            </div>
                            <div className="col-span-1 md:col-span-2 lg:col-span-2 bg-amber-50 dark:bg-amber-900/20 rounded-3xl border border-slate-200 dark:border-amber-800/50 p-5 shadow-sm flex items-center justify-between">
                                <div>
                                    <p className="text-amber-800 dark:text-amber-400 font-semibold mb-1 text-sm">Perlu Pantauan</p>
                                    <h3 className="text-amber-900 dark:text-amber-100 text-3xl font-bold">{problematicProjectsCount} <span className="text-lg font-medium">Beresiko</span></h3>
                                </div>
                                <div className="w-14 h-14 rounded-2xl bg-amber-200/50 dark:bg-amber-800/50 flex items-center justify-center text-amber-700 dark:text-amber-300">
                                    <Icon name="alert-triangle" size={28} />
                                </div>
                            </div>

                            {/* BENTO CELL 5: Tenaga Ahli & Penugasan */}
                            <div className="col-span-1 md:col-span-2 lg:col-span-2 flex gap-4">
                                <div onClick={() => setActiveTab('ahli')} className="flex-1 bg-white/70 dark:bg-slate-800/60 backdrop-blur-2xl rounded-3xl border border-white/60 dark:border-slate-700/50 p-4 shadow-xl shadow-slate-200/50 dark:shadow-slate-900/50 cursor-pointer hover:scale-[1.02] transition-transform flex flex-col justify-center items-center text-center">
                                    <div className="text-blue-600 dark:text-blue-400 mb-2"><Icon name="award" size={28} /></div>
                                    <h4 className="text-2xl font-bold text-slate-800 dark:text-slate-100">{experts.length}</h4>
                                    <p className="text-xs font-medium text-slate-500 mt-1">Tenaga Ahli</p>
                                </div>
                                <div onClick={() => setActiveTab('penugasan')} className="flex-1 bg-white/70 dark:bg-slate-800/60 backdrop-blur-2xl rounded-3xl border border-white/60 dark:border-slate-700/50 p-4 shadow-xl shadow-slate-200/50 dark:shadow-slate-900/50 cursor-pointer hover:scale-[1.02] transition-transform flex flex-col justify-center items-center text-center">
                                    <div className="text-indigo-600 dark:text-indigo-400 mb-2"><Icon name="file-text" size={28} /></div>
                                    <h4 className="text-2xl font-bold text-slate-800 dark:text-slate-100">{assignments.length}</h4>
                                    <p className="text-xs font-medium text-slate-500 mt-1">Penugasan</p>
                                </div>
                            </div>

                            {/* BENTO CELL 6: Logistik & Inventaris */}
                            <div onClick={() => setActiveTab('inventaris')} className="col-span-1 md:col-span-2 lg:col-span-2 bg-slate-800 dark:bg-slate-950 rounded-3xl p-6 shadow-xl shadow-slate-300/50 dark:shadow-none text-white cursor-pointer hover:scale-[1.02] transition-transform flex items-center justify-between overflow-hidden relative group">
                                <div className="absolute right-0 bottom-0 opacity-10 translate-x-2 translate-y-2 group-hover:rotate-12 transition-transform duration-500">
                                    <Icon name="box" size={100} />
                                </div>
                                <div className="relative z-10">
                                    <h4 className="text-4xl font-bold mb-1">{inventory.length}</h4>
                                    <p className="text-sm font-medium text-slate-300">Item Logistik</p>
                                </div>
                                <div className="w-10 h-10 rounded-full border border-slate-600 flex items-center justify-center relative z-10 group-hover:bg-slate-700 transition-colors">
                                    <Icon name="chevron-right" size={18} />
                                </div>
                            </div>
                        </div>

                        <div className="grid grid-cols-1 lg:grid-cols-3 gap-4 lg:gap-8">
                            <div className="lg:col-span-2 glass-card rounded-3xl overflow-hidden flex flex-col">
                                <div className="p-6 border-b border-slate-100 dark:border-slate-700/50 flex justify-between items-center">
                                    <h3 className="text-lg font-bold text-slate-800 dark:text-slate-100">Status Proyek Berjalan</h3>
                                    <button onClick={() => setActiveTab('proyek')} className="text-sm font-medium text-blue-600 hover:text-blue-800">Lihat Semua</button>
                                </div>
                                <div className="p-0 overflow-x-auto min-h-[200px]">
                                    {sortedProjectsForDashboard.length === 0 ? (
                                        <div className="flex flex-col items-center justify-center h-48 text-slate-400">
                                            <Icon name="folder-open" size={32} className="mb-2 opacity-50" />
                                            <p className="text-sm">Belum ada data proyek di Google Sheets.</p>
                                        </div>
                                    ) : (
                                        <table className="enterprise-table text-left">
                                            <thead>
                                                <tr className="bg-slate-50 dark:bg-slate-900/50 text-slate-500 dark:text-slate-400 text-xs uppercase tracking-wider">
                                                    <th className="p-4 font-semibold">Nama Proyek</th>
                                                    <th className="p-4 font-semibold">Progress</th>
                                                    <th className="p-4 font-semibold">Status Makro</th>
                                                </tr>
                                            </thead>
                                            <tbody className="divide-y divide-slate-100 dark:divide-slate-700/50 text-sm [&>tr]:transition-colors [&>tr:hover]:bg-slate-50/50 dark:[&>tr:hover]:bg-slate-800/30">
                                                {/* Menggunakan sortedProjectsForDashboard untuk di-*slice* dan dirender */}
                                                {sortedProjectsForDashboard.slice(0, 5).map((project) => (
                                                    <tr key={project.id} className="hover:bg-slate-50 dark:hover:bg-slate-700/50 transition-colors">
                                                        <td className="p-4">
                                                            <p className="font-semibold text-slate-800 dark:text-slate-200">{project.name}</p>
                                                            <p className="text-xs text-slate-500">{project.client}</p>
                                                            {project.description && <p className="text-[10px] text-slate-400 mt-0.5 line-clamp-1 italic">📝 {project.description}</p>}
                                                        </td>
                                                        <td className="p-4 w-48">
                                                            <div className="flex justify-between text-xs mb-1">
                                                                <span className="font-medium">{project.progress}%</span>
                                                            </div>
                                                            <div className="w-full bg-slate-200 rounded-full h-2">
                                                                <div className={`h-2 rounded-full ${project.computedStatus === 'Terlambat' ? 'bg-red-500' : project.computedStatus === 'Done' ? 'bg-emerald-500' : 'bg-blue-600'}`} style={{ width: `${project.progress}%` }}></div>
                                                            </div>
                                                        </td>
                                                        <td className="p-4"><StatusBadge status={project.computedStatus} /></td>
                                                    </tr>
                                                ))}
                                            </tbody>
                                        </table>
                                    )}
                                </div>
                            </div>

                            <div className="glass-card rounded-3xl p-6 flex flex-col">
                                <div className="flex justify-between items-center mb-2">
                                    <h3 className="text-lg font-bold text-slate-800 dark:text-slate-100">Beban Kerja Tim</h3>
                                    <button onClick={() => setActiveTab('tim')} className="text-sm font-medium text-blue-600 hover:text-blue-800">Lihat Full</button>
                                </div>
                                <p className="text-xs text-slate-500 mb-4 pb-4 border-b border-slate-100 dark:border-slate-700/50">Beban dihitung otomatis dari penugasan.</p>
                                <div className="space-y-5 flex-1 min-h-[150px]">
                                    {calculatedResources.length === 0 ? (
                                        <div className="flex flex-col items-center justify-center h-full text-slate-400">
                                            <Icon name="users" size={32} className="mb-2 opacity-50" />
                                            <p className="text-sm">Belum ada data tim.</p>
                                        </div>
                                    ) : (
                                        calculatedResources.slice(0, 5).map((res) => {
                                            const isOverloaded = res.workload > 100;
                                            const isOptimal = res.workload >= 70 && res.workload <= 100;
                                            return (
                                                <div key={res.id}>
                                                    <div className="flex justify-between items-end mb-1">
                                                        <div>
                                                            <p className="text-sm font-semibold text-slate-800 dark:text-slate-200">{res.name}</p>
                                                            <p className="text-xs text-slate-500 dark:text-slate-400">{res.projects} Proyek Berjalan</p>
                                                        </div>
                                                        <span className={`text-xs font-bold ${isOverloaded ? 'text-red-600' : isOptimal ? 'text-emerald-600' : 'text-blue-600'}`}>{res.workload}%</span>
                                                    </div>
                                                    <div className="w-full bg-slate-100 dark:bg-slate-700/50 rounded-full h-2.5 flex overflow-hidden">
                                                        <div className={`h-full ${isOverloaded ? 'bg-red-500' : isOptimal ? 'bg-emerald-500' : 'bg-blue-500'}`} style={{ width: `${Math.min(res.workload, 100)}%` }}></div>
                                                        {isOverloaded && (<div className="h-full bg-red-800" style={{ width: `${res.workload - 100}%` }}></div>)}
                                                    </div>
                                                </div>
                                            )
                                        })
                                    )}
                                </div>
                            </div>
                        </div>
                    </div>
                );
            };

            // --- KOMPONEN TAB: PROYEK ---
            const renderProyek = () => {
                const getProjectSortScore = (p) => {
                    // 6. Proyek yang sudah selesai (Perencanaan dan Pengawasan)
                    if (p.computedStatus === "Done" || p.status === "Done" || p.progress >= 100) return 60;

                    const isPengawasan = p.type?.toLowerCase().includes('pengawas') || p.type?.toLowerCase().includes('manajemen konstruksi');

                    // 5. Data proyek pengawasan (Running)
                    if (isPengawasan) return 50;

                    // Proyek Perencanaan (Running)
                    // 1. Terlambat dari waktu kontrak
                    if (p.computedStatus === "Terlambat") return 10;

                    // Calculate sub-team statuses
                    let hasMicroTerlambat = false;
                    let hasMicroBeresiko = false;

                    if (p.categoryDetails) {
                        Object.values(p.categoryDetails).forEach(cat => {
                            const microStatus = getMicroStatus(cat.progress || 0, cat.deadline);
                            if (microStatus === "Terlambat") hasMicroTerlambat = true;
                            if (microStatus === "Beresiko") hasMicroBeresiko = true;
                        });
                    }

                    // 2. Macro aman, tapi sub tim banyak yang terlambat
                    if (hasMicroTerlambat) return 20;

                    // 3. Macro aman, sub tim beresiko
                    if (hasMicroBeresiko) return 30;

                    // 4. Macro aman, sub tim aman
                    return 40;
                };

                const filteredAndSortedProjects = computedProjects
                    .filter(p => {
                        const searchLower = searchProjectTab.toLowerCase();
                        const matchesSearch = p.name.toLowerCase().includes(searchLower) ||
                            p.client.toLowerCase().includes(searchLower) ||
                            p.id.toLowerCase().includes(searchLower) ||
                            (p.teamLeader && p.teamLeader.toLowerCase().includes(searchLower)) ||
                            (p.team && p.team.some(member => member.toLowerCase().includes(searchLower))) ||
                            (p.surveyorTeam && p.surveyorTeam.some(member => member.toLowerCase().includes(searchLower)));

                        let matchesType = true;
                        if (filterProjectType !== 'Semua Tipe') {
                            if (filterProjectType === 'Perencanaan') {
                                matchesType = p.type?.toLowerCase().includes('perencana');
                            } else if (filterProjectType === 'Pengawasan') {
                                matchesType = p.type?.toLowerCase().includes('pengawas');
                            } else if (filterProjectType === 'Manajemen Konstruksi') {
                                matchesType = p.type?.toLowerCase().includes('manajemen konstruksi');
                            } else {
                                matchesType = p.type === filterProjectType;
                            }
                        }
                        return matchesSearch && matchesType;
                    })
                    .sort((a, b) => getProjectSortScore(a) - getProjectSortScore(b));

                const totalPages = Math.ceil(filteredAndSortedProjects.length / projectsPerPage);
                const currentProjects = filteredAndSortedProjects.slice(
                    (projectPage - 1) * projectsPerPage,
                    projectPage * projectsPerPage
                );

                return (
                    <div className="space-y-6 fade-in">
                        <ErrorBanner />
                        <div className="glass-card rounded-[2rem] overflow-hidden flex flex-col">
                            <div className="p-6 border-b border-slate-100 dark:border-slate-700/50 flex flex-col md:flex-row md:items-center justify-between gap-4">
                                <div>
                                    <h3 className="text-lg font-bold text-slate-800 dark:text-slate-100">Daftar List Proyek</h3>
                                    <p className="text-sm text-slate-500 dark:text-slate-400">Kelola proyek dan tentukan penugasan timnya di sini.</p>
                                </div>
                                <div className="flex items-center gap-3">
                                    <div className="relative">
                                        <div className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400">
                                            <Icon name="search" size={16} />
                                        </div>
                                        <input
                                            type="text"
                                            placeholder="Cari proyek/klien..."
                                            className="w-full md:w-64 pl-9 pr-4 py-2 bg-slate-50 dark:bg-slate-900/50 border border-slate-200 dark:border-slate-700 rounded-lg text-sm dark:text-slate-200 outline-none focus:ring-2 focus:ring-blue-500 transition-colors placeholder-slate-400 dark:placeholder-slate-500"
                                            value={searchProjectTab}
                                            onChange={(e) => {
                                                setSearchProjectTab(e.target.value);
                                                setProjectPage(1);
                                            }}
                                        />
                                    </div>
                                    <div className="relative">
                                        <select
                                            className="appearance-none w-full md:w-48 pl-3 pr-8 py-2 bg-slate-50 dark:bg-slate-900/50 border border-slate-200 dark:border-slate-700 rounded-lg text-sm text-slate-700 dark:text-slate-300 outline-none focus:ring-2 focus:ring-blue-500 transition-colors cursor-pointer"
                                            value={filterProjectType}
                                            onChange={(e) => {
                                                setFilterProjectType(e.target.value);
                                                setProjectPage(1);
                                            }}
                                        >
                                            <option value="Semua Tipe">Semua Tipe</option>
                                            <option value="Perencanaan">Perencanaan</option>
                                            <option value="Pengawasan">Pengawasan</option>
                                            <option value="Manajemen Konstruksi">Manajemen Konstruksi</option>
                                        </select>
                                        <div className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none">
                                            <Icon name="filter" size={14} />
                                        </div>
                                    </div>
                                    <button onClick={() => openModal('project', 'add')} className="flex items-center gap-2 px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white text-sm font-medium rounded-lg transition-colors shadow-sm whitespace-nowrap">
                                        <Icon name="plus" size={16} /> Tambah
                                    </button>
                                </div>
                            </div>
                            <div className="p-0 overflow-x-auto min-h-[400px]">
                                <table className="enterprise-table text-left">
                                    <thead>
                                        <tr className="bg-slate-50 dark:bg-slate-900/50 text-slate-500 dark:text-slate-400 text-xs uppercase tracking-wider">
                                            <th className="p-4 font-semibold min-w-[200px]">Nama Proyek & Klien</th>
                                            <th className="p-4 font-semibold min-w-[120px]">Target Makro</th>
                                            <th className="p-4 font-semibold min-w-[300px]">Rincian & Target Sub-Tim</th>
                                            <th className="p-4 font-semibold text-right">Aksi</th>
                                        </tr>
                                    </thead>
                                    <tbody className="divide-y divide-slate-100 dark:divide-slate-700/50 text-sm [&>tr]:transition-colors [&>tr:hover]:bg-slate-50/50 dark:[&>tr:hover]:bg-slate-800/30">
                                        {filteredAndSortedProjects.length === 0 ? (
                                            <tr>
                                                <td colSpan="4" className="p-8 text-center text-slate-400">
                                                    <Icon name="folder-open" size={32} className="mx-auto mb-2 opacity-50" />
                                                    <p>Pencarian tidak ditemukan atau data kosong.</p>
                                                </td>
                                            </tr>
                                        ) : (
                                            currentProjects.map((p) => {
                                                const teamGroups = {
                                                    Arsitek: [], QS: [], Struktur: [], MEP: [], 'Tata Ruang': [], Surveyor: [], Lainnya: []
                                                };
                                                const isPerencanaan = p.type?.toLowerCase().includes('perencana');
                                                const surveyorList = p.surveyorTeam || [];

                                                if (p.team && p.team.length > 0) {
                                                    p.team.forEach(teamMemberName => {
                                                        if (isPerencanaan && surveyorList.includes(teamMemberName)) {
                                                            if (!teamGroups.Surveyor.includes(teamMemberName)) {
                                                                teamGroups.Surveyor.push(teamMemberName);
                                                            }
                                                            return; // Memastikan mereka tidak muncul di sub-tim aslinya (hide from default category)
                                                        }

                                                        const resInfo = resources.find(r => r.name === teamMemberName);
                                                        if (resInfo) {
                                                            const cat = getCategoryFromRole(resInfo.role);
                                                            if (teamGroups[cat]) teamGroups[cat].push(resInfo.name);
                                                        } else {
                                                            teamGroups.Lainnya.push(teamMemberName);
                                                        }
                                                    });
                                                }

                                                const renderSubTeamProgress = (title, members) => {
                                                    if (members.length === 0) return null;

                                                    const details = p.categoryDetails?.[title] || {};
                                                    const progress = details.progress || 0;
                                                    const deadline = formatDateIndo(details.deadline);

                                                    const teamWarning = getMicroStatus(progress, details.deadline);

                                                    return (
                                                        <div className="mb-3 last:mb-0 bg-white/80 dark:bg-slate-800/80 backdrop-blur-md border border-white/60 dark:border-slate-700/50 p-2.5 rounded-xl shadow-md shadow-slate-200/40 dark:shadow-none hover:-translate-y-0.5 transition-transform duration-300">
                                                            <div className="flex justify-between items-center text-[11px] mb-1.5 border-b border-slate-100 dark:border-slate-700/50 pb-1.5">
                                                                <div className="flex items-center gap-2">
                                                                    <span className="font-bold text-slate-800 dark:text-slate-200">{title}</span>
                                                                    {teamWarning && teamWarning !== 'On Progress' && teamWarning !== 'Done' && teamWarning !== 'Belum Diatur' && (
                                                                        <span title="Berdasarkan Deadline Sub-Tim" className={`px-1.5 py-0.5 rounded text-[9px] font-bold border flex items-center gap-1 ${teamWarning === 'Terlambat' ? 'bg-red-50 text-red-600 border-red-200' : 'bg-amber-50 text-amber-600 border-amber-200'}`}>
                                                                            <Icon name={teamWarning === 'Terlambat' ? 'alert-triangle' : 'clock'} size={10} />
                                                                            {teamWarning}
                                                                        </span>
                                                                    )}
                                                                </div>
                                                                <span className="text-slate-500 font-medium flex items-center gap-1">
                                                                    <Icon name="calendar-clock" size={10} />
                                                                    {details.deadline ? deadline : 'Belum diatur'}
                                                                </span>
                                                            </div>
                                                            <div className="flex flex-wrap gap-1 mb-2">
                                                                {members.map((name, idx) => (
                                                                    <span key={idx} className="px-1.5 py-0.5 bg-slate-100 dark:bg-slate-700/50 text-slate-600 dark:text-slate-300 rounded text-[10px] font-semibold border border-slate-200 dark:border-slate-600">
                                                                        {name}
                                                                    </span>
                                                                ))}
                                                            </div>
                                                            <div className="flex items-center gap-2">
                                                                <div className="w-full bg-slate-100 dark:bg-slate-700 rounded-full h-1.5 flex overflow-hidden">
                                                                    <div className={`${progress >= 100 ? 'bg-emerald-500' : teamWarning === 'Terlambat' ? 'bg-red-500' : 'bg-blue-500'} h-1.5 rounded-full`} style={{ width: `${progress}%` }}></div>
                                                                </div>
                                                                <span className={`text-[10px] font-bold w-8 text-right ${progress >= 100 ? 'text-emerald-700' : teamWarning === 'Terlambat' ? 'text-red-600' : 'text-blue-700'}`}>{progress}%</span>
                                                            </div>
                                                        </div>
                                                    );
                                                };

                                                return (
                                                    <tr key={p.id} className="hover:bg-indigo-50/50 dark:hover:bg-indigo-900/20 transition-colors group align-top">
                                                        <td className="p-4 pt-5">
                                                            <div className="flex items-center gap-2 mb-0.5">
                                                                <p className="font-bold text-slate-800 dark:text-slate-200 text-base">{p.name}</p>
                                                                {p.notStarted && (
                                                                    <span className="px-2 py-0.5 bg-slate-100 text-slate-700 border border-slate-300 rounded text-[10px] font-bold tracking-wider dark:bg-slate-800 dark:text-slate-300 dark:border-slate-600">BELUM MULAI</span>
                                                                )}
                                                                {p.isPending && (
                                                                    <span className="px-2 py-0.5 bg-orange-100 text-orange-700 border border-orange-200 rounded text-[10px] font-bold tracking-wider">PENDING</span>
                                                                )}
                                                                {p.sourceAssignmentId && (
                                                                    <span title="Proyek ini dikelola dari menu Penugasan Tenaga Ahli" className="inline-flex items-center gap-1 px-2 py-0.5 bg-indigo-50 dark:bg-indigo-900/30 text-indigo-600 dark:text-indigo-400 border border-indigo-200 dark:border-indigo-800/50 rounded text-[10px] font-bold tracking-wider">
                                                                        <Icon name="link" size={10} /> Sync
                                                                    </span>
                                                                )}
                                                            </div>
                                                            <p className="text-xs text-slate-500 mb-2">{p.client}</p>
                                                            {p.isPending && (
                                                                <div className="mb-3 bg-orange-50/50 dark:bg-orange-900/20 border border-orange-200 dark:border-orange-800/50 rounded-lg p-2.5">
                                                                    <div className="flex items-center gap-1.5 mb-1 text-orange-600 dark:text-orange-400">
                                                                        <Icon name="alert-triangle" size={12} />
                                                                        <span className="text-[11px] font-bold">Proyek Sedang Pending</span>
                                                                    </div>
                                                                    <p className="text-[11px] text-slate-600 dark:text-slate-300 italic mb-1">"{p.pendingReason}"</p>
                                                                    {p.pendingDate && (
                                                                        <div className="text-[9px] text-slate-500 flex items-center gap-1 mt-1.5 pt-1.5 border-t border-orange-100 dark:border-orange-900/50">
                                                                            <Icon name="clock" size={10} />
                                                                            Sejak: {p.pendingDate}
                                                                        </div>
                                                                    )}
                                                                </div>
                                                            )}
                                                            <div className="flex flex-wrap gap-1.5 mb-2">
                                                                <span className={`inline-block px-2 py-0.5 rounded text-[10px] font-medium ${p.type?.toLowerCase().includes('perencana') ? 'bg-indigo-50 text-indigo-700 border border-indigo-100' : 'bg-teal-50 text-teal-700 border border-teal-100'}`}>
                                                                    {p.type?.toLowerCase().includes('perencana') ? 'Perencanaan' : p.type?.toLowerCase().includes('pengawas') ? 'Pengawasan' : p.type}
                                                                </span>
                                                                {p.type?.toLowerCase().includes('perencana') && p.divisiKontrol && (
                                                                    <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-medium bg-blue-50 text-blue-700 border border-blue-100">
                                                                        {p.divisiKontrol}
                                                                    </span>
                                                                )}
                                                            </div>
                                                            {p.teamLeader && (
                                                                <div className="mb-2">
                                                                    <span className="inline-flex items-center gap-1.5 px-2 py-1 bg-amber-50 text-amber-700 border border-amber-200 rounded text-xs font-semibold">
                                                                        <Icon name="star" size={12} className="fill-amber-500" />
                                                                        Team Leader: {p.teamLeader}
                                                                    </span>
                                                                </div>
                                                            )}
                                                            {p.description ? (
                                                                <div className="bg-blue-50/50 dark:bg-blue-900/20 rounded border border-blue-100 dark:border-blue-900/50 overflow-hidden">
                                                                    <p className="text-[11px] text-slate-600 dark:text-slate-300 px-2 py-1.5 leading-relaxed whitespace-pre-wrap"><span className="text-blue-400 mr-1">📝</span>{p.description}</p>
                                                                    <p className="text-[9px] text-slate-400 dark:text-slate-500 px-2 pb-1.5 pt-1 mt-1 border-t border-blue-50 dark:border-blue-900/30 bg-blue-50/30 dark:bg-blue-900/10 font-medium">Terakhir diupdate: {p.descriptionUpdatedAt ? formatDateTimeIndo(p.descriptionUpdatedAt) : '-'}</p>
                                                                </div>
                                                            ) : <p className="text-[11px] text-slate-400 italic">Belum ada deskripsi update</p>}
                                                        </td>
                                                        <td className="p-4 pt-5">
                                                            <div className="bg-white/60 dark:bg-slate-900/50 border border-white/50 dark:border-slate-700/50 rounded-xl p-3 shadow-sm">
                                                                <div className="mb-3 border-b border-slate-200 dark:border-slate-700/50 pb-3">
                                                                    <p className="text-[10px] text-slate-500 font-semibold mb-0.5 uppercase tracking-wider">Tgl SPMK Proyek</p>
                                                                    <p className="text-sm font-bold text-slate-800 dark:text-slate-200">{p.spmk ? formatDateIndo(p.spmk) : 'Belum diatur'}</p>
                                                                </div>
                                                                <div className="mb-3">
                                                                    <p className="text-[10px] text-slate-500 font-semibold mb-0.5 uppercase tracking-wider">Tenggat Waktu Kontrak</p>
                                                                    <p className="text-sm font-bold text-slate-800 dark:text-slate-200">{formatDateIndo(p.deadline)}</p>
                                                                </div>
                                                                {!(p.type?.toLowerCase().includes('pengawas') || p.type?.toLowerCase().includes('manajemen konstruksi')) && (
                                                                    <div className="mb-3">
                                                                        <div className="flex items-center justify-between gap-2 text-[10px] font-semibold text-slate-500 mb-1 uppercase tracking-wider">
                                                                            <span className="truncate">Progress</span>
                                                                            <span className="shrink-0">{p.progress}%</span>
                                                                        </div>
                                                                        <div className="w-full bg-slate-200 dark:bg-slate-700 rounded-full h-2">
                                                                            <div className={`h-2 rounded-full ${p.computedStatus === 'Terlambat' ? 'bg-red-500' : p.computedStatus === 'Done' ? 'bg-emerald-500' : 'bg-blue-600'}`} style={{ width: `${p.progress}%` }}></div>
                                                                        </div>
                                                                    </div>
                                                                )}
                                                                <StatusBadge status={p.computedStatus} />
                                                            </div>
                                                        </td>
                                                        <td className="p-4">
                                                            {(!p.team || p.team.length === 0) ? (
                                                                <div className="flex flex-col items-center justify-center h-24 text-slate-400 border border-dashed border-slate-300 dark:border-slate-700 rounded-xl bg-white/40 dark:bg-slate-900/50">
                                                                    <p className="text-xs italic">Belum ada tim dialokasikan</p>
                                                                </div>
                                                            ) : (
                                                                <div className="flex flex-col gap-0 text-xs bg-white/40 dark:bg-slate-900/50 border border-white/50 dark:border-slate-700/50 p-2 rounded-2xl shadow-inner shadow-slate-100/50 dark:shadow-none">
                                                                    {p.type?.toLowerCase().includes('pengawas') || p.type?.toLowerCase().includes('manajemen konstruksi') ? (
                                                                        <div className="space-y-3 mt-1">
                                                                            {['Team Leader', 'Tenaga Ahli', 'Inspector', 'Laboratory Technician', 'Quantity Surveyor', 'K3', 'Administrasi'].map(role => {
                                                                                const roleMembers = p.team.filter(m => (p.pengawasanDetails?.[m]?.role || 'Inspector') === role);
                                                                                if (roleMembers.length === 0) return null;
                                                                                return (
                                                                                    <div key={role} className="flex flex-col mb-3 last:mb-0">
                                                                                        <div className="flex justify-between items-center mb-1.5">
                                                                                            <span className="font-bold text-[11px] text-slate-800 dark:text-slate-200 uppercase tracking-wider">{role}</span>
                                                                                            <span className="text-[10px] px-1.5 py-0.5 rounded bg-emerald-50 dark:bg-emerald-900/30 text-emerald-600 dark:text-emerald-400 font-semibold">{roleMembers.length} Personil</span>
                                                                                        </div>
                                                                                        <div className="space-y-3 pl-2 border-l-2 border-emerald-100 dark:border-emerald-800/50">
                                                                                            {roleMembers.map(m => {
                                                                                                const detail = p.pengawasanDetails?.[m] || {};
                                                                                                const isAutoDone = detail.deadline && detail.deadline < new Date().toISOString().split('T')[0];
                                                                                                const isIndividuallyDone = p.individualStatus?.[m] || isAutoDone;

                                                                                                return (
                                                                                                    <div key={m} className="flex flex-col gap-1 border-b border-slate-100 dark:border-slate-800/50 pb-2 last:border-0 last:pb-0">
                                                                                                        <div className="flex justify-between items-start">
                                                                                                            <span className={`text-[11px] font-semibold ${isIndividuallyDone ? 'text-slate-400 line-through' : 'text-slate-700 dark:text-slate-300'}`}>{m}</span>
                                                                                                            <span className={`text-[9px] px-1.5 py-0.5 rounded font-bold border ${detail.statusTurun === 'Turun' ? 'bg-blue-50 text-blue-600 border-blue-200' : 'bg-slate-50 text-slate-500 border-slate-200'}`}>
                                                                                                                {detail.statusTurun || 'Tidak Turun'}
                                                                                                            </span>
                                                                                                        </div>
                                                                                                        <div className="flex flex-wrap gap-x-3 gap-y-1 text-[9px] text-slate-500">
                                                                                                            <span className="flex items-center gap-1">
                                                                                                                <Icon name="clock" size={10} className="text-slate-400" />
                                                                                                                {detail.manMonth ? `${detail.manMonth} Bulan` : '-'}
                                                                                                            </span>
                                                                                                        </div>
                                                                                                        
                                                                                                    </div>
                                                                                                )
                                                                                            })}
                                                                                        </div>
                                                                                    </div>
                                                                                )
                                                                            })}
                                                                        </div>
                                                                    ) : (
                                                                        <>
                                                                            {renderSubTeamProgress("Arsitek", teamGroups.Arsitek)}
                                                                            {renderSubTeamProgress("Surveyor", teamGroups.Surveyor)}
                                                                            {renderSubTeamProgress("QS", teamGroups.QS)}
                                                                            {renderSubTeamProgress("Struktur", teamGroups.Struktur)}
                                                                            {renderSubTeamProgress("MEP", teamGroups.MEP)}
                                                                            {renderSubTeamProgress("Tata Ruang", teamGroups['Tata Ruang'])}
                                                                            {renderSubTeamProgress("Lainnya", teamGroups.Lainnya)}
                                                                        </>
                                                                    )}
                                                                </div>
                                                            )}
                                                        </td>
                                                        <td className="p-4 text-right lg:opacity-0 lg:group-hover:opacity-100 transition-opacity pt-5">
                                                            <div className="flex justify-end gap-2 flex-col items-end">
                                                                {p.computedStatus === 'Terlambat' && (
                                                                    <button onClick={() => handleAnalyzeDomino(p)} className="flex items-center justify-center gap-1.5 px-3 py-1.5 text-xs font-bold bg-gradient-to-r from-orange-500 to-rose-500 text-white hover:from-orange-600 hover:to-rose-600 rounded-lg shadow-md animate-pulse hover:animate-none"><Icon name="alert-triangle" size={14} /> Cek Domino</button>
                                                                )}

                                                                {p.type?.toLowerCase().includes('perencana') && (
                                                                    <button onClick={() => { setActiveScheduleProject(p); setActiveTab('schedule'); }} className="flex items-center justify-center gap-1.5 px-3 py-1.5 text-xs font-semibold bg-purple-50 dark:bg-purple-900/30 text-purple-600 dark:text-purple-400 hover:bg-purple-100 dark:hover:bg-purple-900/50 rounded-lg border border-purple-200 dark:border-purple-800/50"><Icon name="calendar" size={14} /> Time Schedule</button>
                                                                )}

                                                                {!(p.type?.toLowerCase().includes('pengawas') || p.type?.toLowerCase().includes('manajemen konstruksi')) && (
                                                                    p.isPending ? (
                                                                        <button onClick={() => handleResumeProject(p)} className="flex items-center justify-center gap-1.5 px-3 py-1.5 text-xs font-semibold bg-orange-50 dark:bg-orange-900/30 text-orange-600 dark:text-orange-400 hover:bg-orange-100 dark:hover:bg-orange-900/50 rounded-lg border border-orange-200 dark:border-orange-800/50"><Icon name="play" size={14} /> Resume Proyek</button>
                                                                    ) : (
                                                                        <button onClick={() => { setPendingProjectData(p); setShowPendingModal(true); }} className="flex items-center justify-center gap-1.5 px-3 py-1.5 text-xs font-semibold bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-400 hover:bg-slate-200 dark:hover:bg-slate-700 rounded-lg border border-slate-300 dark:border-slate-600"><Icon name="pause" size={14} /> Set Pending</button>
                                                                    )
                                                                )}

                                                                {p.type?.toLowerCase().includes('pengawas') && (
                                                                    <button onClick={() => handleToggleNotStarted(p.id, !p.notStarted)} className={`flex items-center justify-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg border ${p.notStarted ? 'bg-amber-50 text-amber-600 border-amber-200 hover:bg-amber-100 dark:bg-amber-900/30 dark:border-amber-800/50 dark:hover:bg-amber-900/50' : 'bg-slate-100 text-slate-600 border-slate-300 hover:bg-slate-200 dark:bg-slate-800 dark:border-slate-600 dark:text-slate-400 dark:hover:bg-slate-700'}`}>
                                                                        <Icon name={p.notStarted ? "play" : "pause"} size={14} /> {p.notStarted ? 'Mulai Proyek' : 'Belum Mulai'}
                                                                    </button>
                                                                )}

                                                                <button onClick={() => {
                                                                    if (p.type?.toLowerCase().includes('perencana')) {
                                                                        setPrintZoomProject(p);
                                                                    } else {
                                                                        setPrintData({ type: 'project', id: p.id });
                                                                    }
                                                                }} className="flex items-center justify-center gap-1.5 px-3 py-1.5 text-xs font-semibold bg-indigo-50 dark:bg-indigo-900/30 text-indigo-600 dark:text-indigo-400 hover:bg-indigo-100 dark:hover:bg-indigo-900/50 rounded-lg border border-indigo-200 dark:border-indigo-800/50"><Icon name="printer" size={14} /> Cetak PDF</button>
                                                                <button onClick={() => openModal('project', 'edit', p)} className="flex items-center justify-center gap-1.5 px-3 py-1.5 text-xs font-semibold bg-blue-50 dark:bg-blue-900/30 text-blue-600 dark:text-blue-400 hover:bg-blue-100 dark:hover:bg-blue-900/50 rounded-lg border border-blue-200 dark:border-blue-800/50"><Icon name="edit" size={14} /> Edit Data</button>
                                                                {p.sourceAssignmentId ? (
                                                                    <button onClick={() => setAlertModal({ isOpen: true, title: 'Proyek Tersinkronisasi', message: 'Proyek ini dikelola secara otomatis dari menu Penugasan Tenaga Ahli.\n\nUntuk menghapus proyek ini, silakan hapus data penugasan terkait dari menu Penugasan Tenaga Ahli.' })} className="flex items-center justify-center gap-1.5 px-3 py-1.5 text-xs font-semibold bg-slate-100 dark:bg-slate-800 text-slate-400 dark:text-slate-500 rounded-lg border border-slate-300 dark:border-slate-600 cursor-not-allowed"><Icon name="trash" size={14} /> Hapus</button>
                                                                ) : (
                                                                    <button onClick={() => handleDelete('project', p.id)} className="flex items-center justify-center gap-1.5 px-3 py-1.5 text-xs font-semibold bg-red-50 dark:bg-red-900/30 text-red-600 dark:text-red-400 hover:bg-red-100 dark:hover:bg-red-900/50 rounded-lg border border-red-200 dark:border-red-800/50"><Icon name="trash" size={14} /> Hapus</button>
                                                                )}
                                                            </div>
                                                        </td>
                                                    </tr>
                                                );
                                            })
                                        )}
                                    </tbody>
                                </table>
                            </div>

                            {totalPages > 1 && (
                                <div className="flex items-center justify-between px-4 py-3 border-t border-slate-200 dark:border-slate-700/50 bg-slate-50 dark:bg-slate-900/50 sm:px-6">
                                    <div className="flex justify-between flex-1 sm:hidden">
                                        <button onClick={() => setProjectPage(Math.max(1, projectPage - 1))} disabled={projectPage === 1} className="relative inline-flex items-center px-4 py-2 text-sm font-medium text-slate-700 dark:text-slate-300 bg-white dark:bg-slate-800 border border-slate-300 dark:border-slate-700 rounded-md hover:bg-slate-50 dark:hover:bg-slate-700 disabled:opacity-50">Sebelumnya</button>
                                        <button onClick={() => setProjectPage(Math.min(totalPages, projectPage + 1))} disabled={projectPage === totalPages} className="relative inline-flex items-center px-4 py-2 ml-3 text-sm font-medium text-slate-700 dark:text-slate-300 bg-white dark:bg-slate-800 border border-slate-300 dark:border-slate-700 rounded-md hover:bg-slate-50 dark:hover:bg-slate-700 disabled:opacity-50">Selanjutnya</button>
                                    </div>
                                    <div className="hidden sm:flex sm:flex-1 sm:items-center sm:justify-between">
                                        <div>
                                            <p className="text-xs text-slate-600 dark:text-slate-400">
                                                Menampilkan <span className="font-semibold text-slate-800 dark:text-slate-200">{(projectPage - 1) * projectsPerPage + 1}</span> hingga <span className="font-semibold text-slate-800 dark:text-slate-200">{Math.min(projectPage * projectsPerPage, filteredAndSortedProjects.length)}</span> dari <span className="font-semibold text-slate-800 dark:text-slate-200">{filteredAndSortedProjects.length}</span> proyek
                                            </p>
                                        </div>
                                        <div>
                                            <nav className="relative z-0 inline-flex rounded-md shadow-sm -space-x-px" aria-label="Pagination">
                                                <button onClick={() => setProjectPage(Math.max(1, projectPage - 1))} disabled={projectPage === 1} className="relative inline-flex items-center px-2 py-1.5 rounded-l-md border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-sm font-medium text-slate-500 hover:bg-slate-50 dark:hover:bg-slate-700 disabled:opacity-50">
                                                    <span className="sr-only">Previous</span>
                                                    &larr;
                                                </button>
                                                {Array.from({ length: totalPages }, (_, i) => i + 1).map(num => (
                                                    <button key={num} onClick={() => setProjectPage(num)} className={`relative inline-flex items-center px-3 py-1.5 border text-xs font-semibold ${projectPage === num ? 'z-10 bg-blue-50 border-blue-500 text-blue-600 dark:bg-blue-900/30 dark:border-blue-500/50 dark:text-blue-400' : 'bg-white dark:bg-slate-800 border-slate-300 dark:border-slate-700 text-slate-500 dark:text-slate-400 hover:bg-slate-50 dark:hover:bg-slate-700/50'}`}>
                                                        {num}
                                                    </button>
                                                ))}
                                                <button onClick={() => setProjectPage(Math.min(totalPages, projectPage + 1))} disabled={projectPage === totalPages} className="relative inline-flex items-center px-2 py-1.5 rounded-r-md border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-sm font-medium text-slate-500 hover:bg-slate-50 dark:hover:bg-slate-700 disabled:opacity-50">
                                                    <span className="sr-only">Next</span>
                                                    &rarr;
                                                </button>
                                            </nav>
                                        </div>
                                    </div>
                                </div>
                            )}

                        </div>
                    </div>
                );
            };

            // --- KOMPONEN TAB: TIM & DETAIL PEGAWAI ---
            const renderEmployeeDetail = () => {
                const empCategory = getCategoryFromRole(viewingEmployee.role);

                const activeAssignedProjects = computedProjects.filter(p => {
                    if (!p.team || !p.team.includes(viewingEmployee.name) || p.computedStatus === 'Done' || p.notStarted) return false;
                    const isPengawasan = p.type?.toLowerCase().includes('pengawas') || p.type?.toLowerCase().includes('manajemen konstruksi');
                    let isIndividuallyDone = p.individualStatus?.[viewingEmployee.name] === true;

                    if (isPengawasan) {
                        const pengawasanDeadline = p.pengawasanDetails?.[viewingEmployee.name]?.deadline;
                        if (pengawasanDeadline && pengawasanDeadline < new Date().toISOString().split('T')[0]) {
                            isIndividuallyDone = true;
                        }
                        return !isIndividuallyDone;
                    } else {
                        const effectiveCat = getEffectiveEmpCategory(p, viewingEmployee.name, viewingEmployee.role);
                        const microProgress = p.categoryDetails?.[effectiveCat]?.progress ? Number(p.categoryDetails[effectiveCat].progress) : 0;
                        return microProgress < 100 && !isIndividuallyDone;
                    }
                });

                const completedAssignedProjects = computedProjects.filter(p => {
                    if (!p.team || !p.team.includes(viewingEmployee.name) || p.computedStatus === 'Done' || p.notStarted) return false;
                    const isPengawasan = p.type?.toLowerCase().includes('pengawas') || p.type?.toLowerCase().includes('manajemen konstruksi');
                    let isIndividuallyDone = p.individualStatus?.[viewingEmployee.name] === true;

                    if (isPengawasan) {
                        const pengawasanDeadline = p.pengawasanDetails?.[viewingEmployee.name]?.deadline;
                        if (pengawasanDeadline && pengawasanDeadline < new Date().toISOString().split('T')[0]) {
                            isIndividuallyDone = true;
                        }
                        return isIndividuallyDone;
                    } else {
                        const effectiveCat = getEffectiveEmpCategory(p, viewingEmployee.name, viewingEmployee.role);
                        const microProgress = p.categoryDetails?.[effectiveCat]?.progress ? Number(p.categoryDetails[effectiveCat].progress) : 0;
                        return microProgress >= 100 || isIndividuallyDone;
                    }
                });

                const ledProjects = computedProjects.filter(p =>
                    p.teamLeader === viewingEmployee.name &&
                    p.computedStatus !== 'Done' && !p.notStarted
                );

                return (
                    <div className="space-y-6 fade-in">
                        <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
                            <div className="flex items-center gap-4">
                                <button onClick={() => setViewingEmployee(null)} className="p-2.5 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700/50 rounded-xl hover:bg-slate-50 dark:hover:bg-slate-700/50 transition-colors shadow-sm">
                                    <Icon name="arrow-left" size={20} className="text-slate-600 dark:text-slate-300" />
                                </button>
                                <div>
                                    <h2 className="text-lg sm:text-xl lg:text-2xl font-bold text-slate-800 dark:text-slate-100">Rincian Penugasan: {viewingEmployee.name}</h2>
                                    {viewingEmployee.level && viewingEmployee.level.startsWith('Kordinator Divisi') ? (
                                        <div className="text-sm text-slate-500">
                                            <p>{viewingEmployee.level}</p>
                                            <p>{viewingEmployee.role} • Kapasitas Terpakai: <strong className={viewingEmployee.workload > 100 ? 'text-red-600' : 'text-blue-600'}>{viewingEmployee.workload}%</strong></p>
                                        </div>
                                    ) : (
                                        <p className="text-sm text-slate-500">
                                            {viewingEmployee.level === 'PIC' ? `PIC ${viewingEmployee.role}` : viewingEmployee.role} • Kapasitas Terpakai: <strong className={viewingEmployee.workload > 100 ? 'text-red-600' : 'text-blue-600'}>{viewingEmployee.workload}%</strong>
                                        </p>
                                    )}
                                </div>
                            </div>
                            <button onClick={() => setPrintData({ type: 'personnel', name: viewingEmployee.name })} className="flex items-center gap-2 px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-medium rounded-lg transition-colors shadow-sm shrink-0">
                                <Icon name="printer" size={16} /> Cetak Laporan Pegawai
                            </button>
                        </div>

                        <div className="glass-card rounded-[2rem] overflow-hidden flex flex-col">
                            {viewingEmployee.level === 'Team Leader' && (
                                <div className="mb-6 border-b border-slate-200 dark:border-slate-700/50">
                                    <div className="p-6 bg-amber-50 dark:bg-amber-900/20 border-b border-amber-200 dark:border-amber-800/50 flex items-center gap-2">
                                        <Icon name="star" size={20} className="fill-amber-500 text-amber-600" />
                                        <div>
                                            <h3 className="text-lg font-bold text-amber-900 dark:text-amber-100">Daftar Proyek yang Di-Lead</h3>
                                            <p className="text-sm text-amber-700 dark:text-amber-300">Tanggung jawab pengawasan keseluruhan proyek secara Makro.</p>
                                        </div>
                                    </div>
                                    <div className="p-0 overflow-x-auto min-h-[150px]">
                                        <table className="enterprise-table text-left">
                                            <thead>
                                                <tr className="bg-amber-50/50 dark:bg-amber-900/10 text-amber-800 dark:text-amber-300 text-xs uppercase tracking-wider border-b border-slate-200 dark:border-amber-800/50">
                                                    <th className="p-4 font-semibold">Nama Proyek & Klien</th>
                                                    <th className="p-4 font-semibold text-center">Tenggat Waktu Kontrak</th>
                                                    <th className="p-4 font-semibold">Total Progress</th>
                                                    <th className="p-4 font-semibold">Status Makro</th>
                                                </tr>
                                            </thead>
                                            <tbody className="divide-y divide-amber-100 dark:divide-amber-800/50 text-sm bg-white dark:bg-slate-800">
                                                {ledProjects.length === 0 ? (
                                                    <tr>
                                                        <td colSpan="4" className="p-8 text-center text-slate-400">
                                                            <Icon name="briefcase" size={32} className="mx-auto mb-2 opacity-50" />
                                                            <p>Belum ada proyek yang dipimpin.</p>
                                                        </td>
                                                    </tr>
                                                ) : (
                                                    ledProjects.map(p => (
                                                        <tr key={`led-${p.id}`} className="hover:bg-amber-50/30 dark:hover:bg-amber-900/10 transition-colors">
                                                            <td className="p-4">
                                                                <p className="font-semibold text-slate-800 dark:text-slate-200">{p.name}</p>
                                                                <p className="text-xs text-slate-500">{p.client}</p>
                                                            </td>
                                                            <td className="p-4 text-center font-medium text-slate-700 dark:text-slate-300">
                                                                {p.deadline ? formatDateIndo(p.deadline) : <span className="text-slate-400 italic text-xs">Belum diatur</span>}
                                                            </td>
                                                            <td className="p-4 w-48">
                                                                <div className="flex justify-between text-xs mb-1">
                                                                    <span className="font-medium text-slate-700 dark:text-slate-300">{p.progress}% Selesai</span>
                                                                </div>
                                                                <div className="w-full bg-slate-200 dark:bg-slate-700 rounded-full h-2">
                                                                    <div className={`h-2 rounded-full ${p.computedStatus === 'Terlambat' ? 'bg-red-500' : p.computedStatus === 'Done' ? 'bg-emerald-500' : 'bg-blue-600'}`} style={{ width: `${p.progress}%` }}></div>
                                                                </div>
                                                            </td>
                                                            <td className="p-4">
                                                                <StatusBadge status={p.computedStatus} />
                                                            </td>
                                                        </tr>
                                                    ))
                                                )}
                                            </tbody>
                                        </table>
                                    </div>
                                </div>
                            )}

                            <div className="p-6 border-b border-slate-100 dark:border-slate-700/50">
                                <h3 className="text-lg font-bold text-slate-800 dark:text-slate-100">Daftar Penugasan Sub-Tim Aktif</h3>
                                <p className="text-sm text-slate-500">Menampilkan target progres dan deadline spesifik untuk Sub-Tim <strong>{empCategory}</strong> pada masing-masing proyek.</p>
                            </div>
                            <div className="p-0 overflow-x-auto min-h-[300px]">
                                <table className="enterprise-table text-left">
                                    <thead>
                                        <tr className="bg-slate-50 dark:bg-slate-900/50 text-slate-500 dark:text-slate-400 text-xs uppercase tracking-wider">
                                            <th className="p-4 font-semibold">Nama Proyek</th>
                                            <th className="p-4 font-semibold">Status Proyek (Makro)</th>
                                            <th className="p-4 font-semibold">Tugas Sub-Tim ({empCategory})</th>
                                            <th className="p-4 font-semibold">Deadline Sub-Tim</th>
                                            <th className="p-4 font-semibold text-center">Status Individu</th>
                                        </tr>
                                    </thead>
                                    <tbody className="divide-y divide-slate-100 dark:divide-slate-700/50 text-sm [&>tr]:transition-colors [&>tr:hover]:bg-slate-50/50 dark:[&>tr:hover]:bg-slate-800/30">
                                        {activeAssignedProjects.length === 0 ? (
                                            <tr>
                                                <td colSpan="5" className="p-8 text-center text-slate-400">
                                                    <Icon name="briefcase" size={32} className="mx-auto mb-2 opacity-50" />
                                                    <p>Belum ada proyek yang ditugaskan kepada personil ini.</p>
                                                </td>
                                            </tr>
                                        ) : (
                                            activeAssignedProjects.map(p => {
                                                const isPengawasan = p.type?.toLowerCase().includes('pengawas') || p.type?.toLowerCase().includes('manajemen konstruksi');
                                                const pengawasanDetail = p.pengawasanDetails?.[viewingEmployee.name] || {};
                                                const mRole = pengawasanDetail.role || '-';

                                                const effectiveCat = getEffectiveEmpCategory(p, viewingEmployee.name, viewingEmployee.role);
                                                const microDetails = p.categoryDetails?.[effectiveCat] || {};
                                                const mProg = microDetails.progress || 0;
                                                const mDead = isPengawasan ? pengawasanDetail.deadline : microDetails.deadline;
                                                const isIndividuallyDone = p.individualStatus && p.individualStatus[viewingEmployee.name] === true;
                                                const mStatus = isIndividuallyDone || mProg === 100 ? 'Done' : getMicroStatus(mProg, mDead);

                                                return (
                                                    <tr key={p.id} className="hover:bg-indigo-50/50 dark:hover:bg-indigo-900/20 transition-colors">
                                                        <td className="p-4">
                                                            <p className="font-semibold text-slate-800 dark:text-slate-200">{p.name}</p>
                                                            <p className="text-xs text-slate-500">{p.client}</p>
                                                        </td>
                                                        <td className="p-4">
                                                            <StatusBadge status={p.computedStatus} />
                                                        </td>
                                                        {isPengawasan ? (
                                                            <td colSpan="2" className="p-4">
                                                                <div className="flex flex-col gap-1">
                                                                    <span className="text-[11px] font-bold text-slate-700 dark:text-slate-300 uppercase tracking-wider">Peran: {mRole} | Status: {pengawasanDetail.statusTurun || 'Tidak Turun'}</span>
                                                                    <span className="text-[11px] text-slate-600 dark:text-slate-400">SPMK: {p.spmk ? formatDateIndo(p.spmk) : '-'} | {pengawasanDetail.manMonth || '-'} Bulan</span>
                                                                    
                                                                </div>
                                                            </td>
                                                        ) : (
                                                            <>
                                                                <td className="p-4 w-48">
                                                                    <div className="flex justify-between text-xs mb-1">
                                                                        <span className="font-medium text-slate-700 dark:text-slate-300">{mProg}% Selesai {isIndividuallyDone && <span className="text-emerald-600 font-bold ml-1">(Bebas)</span>}</span>
                                                                    </div>
                                                                    <div className="w-full bg-slate-200 dark:bg-slate-700 rounded-full h-2">
                                                                        <div className={`h-2 rounded-full ${mStatus === 'Terlambat' ? 'bg-red-500' : mStatus === 'Done' ? 'bg-emerald-500' : 'bg-blue-600'}`} style={{ width: `${mProg}%` }}></div>
                                                                    </div>
                                                                </td>
                                                                <td className="p-4 font-medium text-slate-700 dark:text-slate-300">
                                                                    {mDead ? formatDateIndo(mDead) : <span className="text-slate-400 italic text-xs">Belum diatur</span>}
                                                                </td>
                                                            </>
                                                        )}
                                                        <td className="p-4 text-center">
                                                            <button
                                                                className={`btn-clean-check group ${isIndividuallyDone ? 'success' : ''}`}
                                                                onClick={() => handleToggleIndividualStatus(p.id, viewingEmployee.name, !isIndividuallyDone)}
                                                                title={isIndividuallyDone ? 'Batalkan Selesai' : 'Tandai Selesai'}
                                                            >
                                                                {!isIndividuallyDone ? (
                                                                    <Icon name="check" size={20} className="transition-colors" />
                                                                ) : (
                                                                    <div className="flip-wrapper">
                                                                        <div className="flip-inner">
                                                                            <div className="flip-front">
                                                                                <Icon name="check" size={20} className="transition-colors" />
                                                                            </div>
                                                                            <div className="flip-back">
                                                                                <Icon name="x" size={20} className="transition-colors" />
                                                                            </div>
                                                                        </div>
                                                                    </div>
                                                                )}
                                                            </button>
                                                        </td>
                                                    </tr>
                                                )
                                            })
                                        )}
                                    </tbody>
                                </table>
                            </div>

                            {completedAssignedProjects.length > 0 && (
                                <>
                                    <div className="p-6 border-b border-slate-100 dark:border-slate-700/50 mt-4">
                                        <h3 className="text-lg font-bold text-slate-800 dark:text-slate-100">Riwayat Penugasan Selesai</h3>
                                        <p className="text-sm text-slate-500">Menampilkan proyek di mana tugas Sub-Tim <strong>{empCategory}</strong> telah mencapai 100% atau ditandai selesai secara individu. Proyek ini tidak lagi membebani kapasitas kerja personil.</p>
                                    </div>
                                    <div className="p-0 overflow-x-auto min-h-[150px]">
                                        <table className="enterprise-table text-left opacity-80">
                                            <thead>
                                                <tr className="bg-emerald-50/50 dark:bg-emerald-900/20 text-emerald-800 dark:text-emerald-300 text-xs uppercase tracking-wider">
                                                    <th className="p-4 font-semibold">Nama Proyek</th>
                                                    <th className="p-4 font-semibold">Status Proyek (Makro)</th>
                                                    <th className="p-4 font-semibold">Tugas Sub-Tim ({empCategory})</th>
                                                    <th className="p-4 font-semibold">Deadline Sub-Tim</th>
                                                    <th className="p-4 font-semibold text-center">Status Individu</th>
                                                </tr>
                                            </thead>
                                            <tbody className="divide-y divide-emerald-50 dark:divide-emerald-900/20 text-sm">
                                                {completedAssignedProjects.map(p => {
                                                    const isPengawasan = p.type?.toLowerCase().includes('pengawas') || p.type?.toLowerCase().includes('manajemen konstruksi');
                                                    const pengawasanDetail = p.pengawasanDetails?.[viewingEmployee.name] || {};
                                                    const mRole = pengawasanDetail.role || '-';

                                                    const effectiveCat = getEffectiveEmpCategory(p, viewingEmployee.name, viewingEmployee.role);
                                                    const microDetails = p.categoryDetails?.[effectiveCat] || {};
                                                    const mProg = microDetails.progress || 0;
                                                    const mDead = isPengawasan ? pengawasanDetail.deadline : microDetails.deadline;

                                                    let isIndividuallyDone = p.individualStatus?.[viewingEmployee.name] === true;
                                                    const pengawasanDeadline = isPengawasan ? p.pengawasanDetails?.[viewingEmployee.name]?.deadline : null;
                                                    const isAutoDone = isPengawasan && pengawasanDeadline && pengawasanDeadline < new Date().toISOString().split('T')[0];

                                                    if (isAutoDone) isIndividuallyDone = true;

                                                    return (
                                                        <tr key={p.id} className="hover:bg-emerald-50/30 dark:hover:bg-emerald-900/10 transition-colors">
                                                            <td className="p-4">
                                                                <p className="font-semibold text-slate-800 dark:text-slate-200">{p.name}</p>
                                                                <p className="text-xs text-slate-500">{p.client}</p>
                                                            </td>
                                                            <td className="p-4">
                                                                <StatusBadge status={p.computedStatus} />
                                                            </td>
                                                            {isPengawasan ? (
                                                                <td colSpan="2" className="p-4">
                                                                    <div className="flex flex-col gap-1">
                                                                        <span className="text-[11px] font-bold text-slate-700 dark:text-slate-300 uppercase tracking-wider">Peran: {mRole} | Status: {pengawasanDetail.statusTurun || 'Tidak Turun'}</span>
                                                                        <span className="text-[11px] text-slate-600 dark:text-slate-400">SPMK: {p.spmk ? formatDateIndo(p.spmk) : '-'} | {pengawasanDetail.manMonth || '-'} Bulan</span>
                                                                        
                                                                    </div>
                                                                </td>
                                                            ) : (
                                                                <>
                                                                    <td className="p-4 w-48">
                                                                        <div className="flex justify-between text-xs mb-1">
                                                                            <span className="font-medium text-slate-700 dark:text-slate-300">{mProg}% Selesai {isIndividuallyDone && <span className="text-emerald-600 font-bold ml-1">(Bebas)</span>}</span>
                                                                        </div>
                                                                        <div className="w-full bg-slate-200 dark:bg-slate-700 rounded-full h-2">
                                                                            <div className={`h-2 rounded-full bg-emerald-500`} style={{ width: `${mProg}%` }}></div>
                                                                        </div>
                                                                    </td>
                                                                    <td className="p-4 font-medium text-slate-700 dark:text-slate-300">
                                                                        {mDead ? formatDateIndo(mDead) : <span className="text-slate-400 italic text-xs">Belum diatur</span>}
                                                                    </td>
                                                                </>
                                                            )}
                                                            <td className="p-4 text-center">
                                                                {(mProg >= 100 && !isIndividuallyDone) || isAutoDone ? (
                                                                    <div className="px-3 py-2 rounded-xl text-xs font-bold border flex items-center justify-center gap-1.5 w-full bg-slate-50/50 dark:bg-slate-800/50 text-slate-500 dark:text-slate-400 border-slate-200 dark:border-slate-700 cursor-default shadow-inner" title="Telah diselesaikan secara otomatis">
                                                                        <Icon name="check-circle" size={14} className="text-emerald-500/70" />
                                                                        Selesai (Otomatis)
                                                                    </div>
                                                                ) : (
                                                                    <button
                                                                        className={`btn-clean-check group ${isIndividuallyDone ? 'success' : ''}`}
                                                                        onClick={() => handleToggleIndividualStatus(p.id, viewingEmployee.name, !isIndividuallyDone)}
                                                                        title={isIndividuallyDone ? 'Batalkan Selesai' : 'Tandai Selesai'}
                                                                    >
                                                                        {!isIndividuallyDone ? (
                                                                            <Icon name="check" size={20} className="transition-colors" />
                                                                        ) : (
                                                                            <div className="flip-wrapper">
                                                                                <div className="flip-inner">
                                                                                    <div className="flip-front">
                                                                                        <Icon name="check" size={20} className="transition-colors" />
                                                                                    </div>
                                                                                    <div className="flip-back">
                                                                                        <Icon name="x" size={20} className="transition-colors" />
                                                                                    </div>
                                                                                </div>
                                                                            </div>
                                                                        )}
                                                                    </button>
                                                                )}
                                                            </td>
                                                        </tr>
                                                    )
                                                })}
                                            </tbody>
                                        </table>
                                    </div>
                                </>
                            )}
                        </div>
                    </div>
                );
            };

            const renderGantt = () => {
                const today = new Date();
                today.setHours(0, 0, 0, 0);

                const ganttData = resources.map(res => {
                    let latestDeadlineStr = null;
                    let latestDate = new Date(0);
                    let latestProjectName = null;

                    computedProjects.forEach(p => {
                        if (p.computedStatus === 'Done') return; // Hanya yang aktif

                        const isPengawasan = p.type?.toLowerCase().includes('pengawas') || p.type?.toLowerCase().includes('manajemen konstruksi');
                        if (isPengawasan) return; // Abaikan proyek pengawasan di Gantt

                        let projDeadlineStr = null;

                        // Khusus untuk Team Leader, abaikan deadline utama proyek. 
                        // Hanya ambil deadline dari perannya sebagai anggota Sub-Tim.
                        if (p.team && p.team.includes(res.name)) {
                            const effectiveCat = getEffectiveEmpCategory(p, res.name, res.role);
                            const isIndividuallyDone = p.individualStatus?.[res.name] === true;
                            if (!isIndividuallyDone && p.categoryDetails?.[effectiveCat]?.deadline) {
                                projDeadlineStr = p.categoryDetails[effectiveCat].deadline;
                            }
                        }

                        if (projDeadlineStr) {
                            const d = new Date(projDeadlineStr);
                            if (!isNaN(d.getTime()) && d > latestDate) {
                                latestDate = d;
                                latestDeadlineStr = projDeadlineStr;
                                latestProjectName = p.name;
                            }
                        }
                    });

                    let availableNow = false;
                    let isOverdue = false;
                    let remainingDays = 0;

                    if (!latestDeadlineStr) {
                        availableNow = true;
                    } else if (latestDate < today) {
                        isOverdue = true;
                        const diffTime = today - latestDate;
                        remainingDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
                    } else {
                        const diffTime = latestDate - today;
                        remainingDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
                    }

                    return {
                        ...res,
                        latestDeadlineStr,
                        latestDate,
                        latestProjectName,
                        availableNow,
                        isOverdue,
                        remainingDays
                    };
                });

                ganttData.sort((a, b) => {
                    if (a.isOverdue && !b.isOverdue) return -1;
                    if (!a.isOverdue && b.isOverdue) return 1;
                    if (a.availableNow !== b.availableNow) return a.availableNow ? 1 : -1;
                    if (a.isOverdue) return b.remainingDays - a.remainingDays; // Sort highest overdue first
                    return a.remainingDays - b.remainingDays;
                });

                const maxDays = Math.max(...ganttData.map(d => d.remainingDays), 30);

                return (
                    <div className="space-y-6 fade-in">
                        <ErrorBanner />
                        <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 glass-card p-4 sm:p-6 rounded-2xl sm:rounded-3xl">
                            <div>
                                <h3 className="text-lg font-bold text-slate-800 dark:text-slate-100">Ploting & Jadwal Personil Perencanaan</h3>
                                <p className="text-sm text-slate-500 dark:text-slate-400 mt-1 leading-relaxed">Visualisasi sisa waktu beban kerja setiap personil dihitung dari <strong className="text-slate-700 dark:text-slate-300">Hari Ini</strong> hingga target <em className="text-slate-700 dark:text-slate-300">Deadline</em> terjauh dari semua proyek aktif mereka. Sangat berguna untuk memutuskan ploting personil ke proyek baru.</p>
                            </div>
                            <div className="w-full sm:w-64 flex-shrink-0 relative">
                                <Icon name="search" size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                                <input
                                    type="text"
                                    placeholder="Cari pegawai, posisi, tugas..."
                                    value={searchGanttTab}
                                    onChange={e => setSearchGanttTab(e.target.value)}
                                    className="w-full pl-9 pr-4 py-2 bg-slate-50 dark:bg-slate-900/50 border border-slate-200 dark:border-slate-700 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 dark:focus:border-blue-400 transition-all dark:text-slate-200"
                                />
                            </div>
                        </div>

                        <div className="bg-white/70 dark:bg-slate-800/60 backdrop-blur-2xl rounded-2xl sm:rounded-3xl shadow-xl shadow-slate-200/50 dark:shadow-slate-900/50 border border-white/60 dark:border-slate-700/50 overflow-hidden">
                            <div className="overflow-x-auto custom-scrollbar">
                                <div className="min-w-[600px] p-4 sm:p-6 space-y-5">
                                    {(() => {
                                        let filteredGanttData = ganttData;
                                        if (searchGanttTab.trim()) {
                                            const term = searchGanttTab.toLowerCase();
                                            filteredGanttData = ganttData.filter(d =>
                                                d.name.toLowerCase().includes(term) ||
                                                d.role.toLowerCase().includes(term) ||
                                                (d.latestProjectName && d.latestProjectName.toLowerCase().includes(term))
                                            );
                                        }

                                        if (filteredGanttData.length === 0) {
                                            return (
                                                <div className="text-center py-8">
                                                    <div className="inline-flex items-center justify-center w-12 h-12 rounded-full bg-slate-100 dark:bg-slate-700/50 text-slate-400 dark:text-slate-500 mb-3">
                                                        <Icon name="search" size={24} />
                                                    </div>
                                                    <h3 className="text-base font-bold text-slate-700 dark:text-slate-300">Tidak ada data ditemukan</h3>
                                                    <p className="text-sm text-slate-500 mt-1">Coba gunakan kata kunci pencarian yang lain.</p>
                                                </div>
                                            );
                                        }

                                        return filteredGanttData.map((data, idx) => {
                                            let widthPct = 0;
                                            let barColor = "bg-slate-200";
                                            let textColor = "text-slate-500";
                                            let label = "Siap Ditugaskan Saat Ini";

                                            if (data.isOverdue) {
                                                widthPct = 100;
                                                label = `Terlambat dari ${formatDateIndo(data.latestDeadlineStr)}`;
                                                barColor = "bg-rose-600";
                                                textColor = "text-rose-700";
                                            } else if (!data.availableNow) {
                                                widthPct = Math.min((data.remainingDays / maxDays) * 100, 100);
                                                label = `s/d ${formatDateIndo(data.latestDeadlineStr)}`;

                                                if (data.remainingDays <= 14) { barColor = "bg-emerald-400"; textColor = "text-emerald-700"; }
                                                else if (data.remainingDays <= 60) { barColor = "bg-blue-500"; textColor = "text-blue-700"; }
                                                else { barColor = "bg-indigo-500"; textColor = "text-indigo-700"; }
                                            } else {
                                                barColor = "bg-emerald-500"; textColor = "text-emerald-700";
                                            }

                                            return (
                                                <div key={idx} className="flex flex-col gap-1.5 border-b border-slate-50 dark:border-slate-700/50 pb-4 last:border-0 last:pb-0">
                                                    <div className="flex justify-between items-end mb-1">
                                                        <div className="flex flex-col gap-0.5">
                                                            <span className="text-sm font-bold text-slate-800 dark:text-slate-200 flex flex-wrap items-center gap-2">
                                                                {data.name}
                                                                <span className="text-[10px] font-bold text-slate-500 dark:text-slate-400 bg-slate-100 dark:bg-slate-700 px-2.5 py-0.5 rounded-full border border-slate-200 dark:border-slate-600">{data.role}</span>
                                                            </span>
                                                            {data.latestProjectName && !data.availableNow && (
                                                                <span className="text-[10px] font-medium text-slate-400">
                                                                    Tugas Terakhir: {data.latestProjectName}
                                                                </span>
                                                            )}
                                                        </div>
                                                        <span className={`text-[11px] font-bold whitespace-nowrap ml-2 ${textColor}`}>{label}</span>
                                                    </div>
                                                    <div className="w-full bg-slate-100 dark:bg-slate-700/50 h-7 rounded-full overflow-hidden relative shadow-inner flex border border-slate-200/50 dark:border-slate-600/50">
                                                        {data.availableNow ? (
                                                            <div className={`h-full ${barColor} w-full flex items-center px-4 justify-between bg-[url('data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSI0IiBoZWlnaHQ9IjQiPjxyZWN0IHdpZHRoPSI0IiBoZWlnaHQ9IjQiIGZpbGw9IiNmZmYiIGZpbGwtb3BhY2l0eT0iMC4xIi8+PC9zdmc+')]`}>
                                                                <span className="text-[11px] font-bold text-white drop-shadow-sm flex items-center gap-1.5"><Icon name="check-circle-2" size={14} /> Available Now</span>
                                                            </div>
                                                        ) : data.isOverdue ? (
                                                            <div className={`h-full ${barColor} w-full flex items-center px-4 justify-between bg-[url('data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSI0IiBoZWlnaHQ9IjQiPjxyZWN0IHdpZHRoPSI0IiBoZWlnaHQ9IjQiIGZpbGw9IiNmZmYiIGZpbGwtb3BhY2l0eT0iMC4xIi8+PC9zdmc+')]`}>
                                                                <span className="text-[11px] font-bold text-white drop-shadow-sm flex items-center gap-1.5"><Icon name="alert-triangle" size={14} /> Terlambat {data.remainingDays} Hari!</span>
                                                            </div>
                                                        ) : (
                                                            <div className={`h-full ${barColor} flex items-center px-3 min-w-[5.5rem] justify-end rounded-full bg-gradient-to-r from-transparent to-black/10`} style={{ width: `${Math.max(widthPct, 12)}%` }}>
                                                                <span className="text-[11px] font-bold text-white drop-shadow-sm whitespace-nowrap">{data.remainingDays} Hari Lagi</span>
                                                            </div>
                                                        )}
                                                    </div>
                                                </div>
                                            );
                                        })
                                    })()}
                                </div>
                            </div>
                        </div>
                    </div>
                );
            };

            const renderTim = () => {
                if (viewingEmployee) return renderEmployeeDetail();

                const filteredResourcesTab = calculatedResources.filter(res =>
                    res.name.toLowerCase().includes(searchTeamTab.toLowerCase()) ||
                    res.role.toLowerCase().includes(searchTeamTab.toLowerCase()) ||
                    (res.level && res.level.toLowerCase().includes(searchTeamTab.toLowerCase()))
                );

                return (
                    <div className="space-y-6 fade-in">
                        <ErrorBanner />
                        <div className="glass-card rounded-[2rem] overflow-hidden flex flex-col">
                            <div className="p-6 border-b border-slate-100 dark:border-slate-700/50 flex flex-col md:flex-row md:items-center justify-between gap-4 bg-white/40 dark:bg-slate-800/40">
                                <div>
                                    <h3 className="text-lg font-bold text-slate-800 dark:text-slate-100">Alokasi & Beban Kerja Tim</h3>
                                    <p className="text-sm text-slate-500">Otomatis dihitung dari jumlah proyek yang mereka pegang (1 Proyek = 25% Beban).</p>
                                </div>
                                <div className="flex items-center gap-3">
                                    <div className="relative">
                                        <div className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400">
                                            <Icon name="search" size={16} />
                                        </div>
                                        <input
                                            type="text"
                                            placeholder="Cari nama/Tim..."
                                            className="w-full md:w-64 pl-9 pr-4 py-2 bg-slate-50 dark:bg-slate-900/50 border border-slate-200 dark:border-slate-700 rounded-xl text-sm outline-none focus:ring-2 focus:ring-blue-500 transition-colors dark:text-slate-200"
                                            value={searchTeamTab}
                                            onChange={(e) => setSearchTeamTab(e.target.value)}
                                        />
                                    </div>
                                    <button onClick={() => openModal('team', 'add')} className="flex items-center gap-2 px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-medium rounded-lg transition-colors shadow-sm whitespace-nowrap">
                                        <Icon name="plus" size={16} /> Tambah
                                    </button>
                                </div>
                            </div>
                            <div className="p-0 overflow-x-auto min-h-[400px]">
                                <table className="enterprise-table text-left">
                                    <thead>
                                        <tr className="bg-slate-50 dark:bg-slate-900/50 text-slate-500 dark:text-slate-400 text-xs uppercase tracking-wider">
                                            <th className="p-4 font-semibold whitespace-nowrap">Nama & Tim</th>
                                            <th className="p-4 font-semibold whitespace-nowrap text-center">Total Proyek</th>
                                            <th className="p-4 font-semibold whitespace-nowrap">Indikator Beban Kerja</th>
                                            <th className="p-4 font-semibold whitespace-nowrap text-right">Aksi</th>
                                        </tr>
                                    </thead>
                                    <tbody className="divide-y divide-slate-100 dark:divide-slate-700/50 text-sm [&>tr]:transition-colors [&>tr:hover]:bg-slate-50/50 dark:[&>tr:hover]:bg-slate-800/30">
                                        {filteredResourcesTab.length === 0 ? (
                                            <tr>
                                                <td colSpan="4" className="p-8 text-center text-slate-400">
                                                    <Icon name="users" size={32} className="mx-auto mb-2 opacity-50" />
                                                    <p>Pencarian tidak ditemukan.</p>
                                                </td>
                                            </tr>
                                        ) : (
                                            filteredResourcesTab.map((res) => {
                                                const isOverloaded = res.workload > 100;
                                                return (
                                                    <tr key={res.id} className="hover:bg-indigo-50/50 dark:hover:bg-indigo-900/20 transition-colors group">
                                                        <td className="p-4">
                                                            <p className="font-semibold text-slate-800 dark:text-slate-200">{res.name}</p>
                                                            {res.level && res.level.startsWith('Kordinator Divisi') ? (
                                                                <div className="text-xs text-slate-500 dark:text-slate-400">
                                                                    <p>{res.level}</p>
                                                                    <p>{res.role}</p>
                                                                </div>
                                                            ) : (
                                                                <p className="text-xs text-slate-500 dark:text-slate-400">
                                                                    {res.level === 'PIC' ? `PIC ${res.role}` : res.role}
                                                                </p>
                                                            )}
                                                            {res.level === 'Team Leader' && (
                                                                <p className="text-[10px] font-bold text-amber-600 mt-1 flex items-center gap-1"><Icon name="star" size={10} className="fill-amber-500" /> Team Leader</p>
                                                            )}
                                                        </td>
                                                        <td className="p-4 text-center"><span className="inline-block px-3 py-1 bg-slate-100 dark:bg-slate-700 text-slate-700 dark:text-slate-300 rounded-full text-xs font-semibold whitespace-nowrap">{`${res.projects} Proyek Aktif`}</span></td>
                                                        <td className="p-4 w-64">
                                                            <div className="flex justify-between items-end mb-1">
                                                                <span className={`text-xs font-bold ${isOverloaded ? 'text-red-600' : 'text-blue-600'}`}>{res.workload}% Kapasitas Terpakai</span>
                                                            </div>
                                                            <div className="w-full bg-slate-100 dark:bg-slate-700 rounded-full h-2.5 flex overflow-hidden">
                                                                <div className={`h-full ${isOverloaded ? 'bg-red-500' : 'bg-blue-500'}`} style={{ width: `${Math.min(res.workload, 100)}%` }}></div>
                                                                {isOverloaded && (<div className="h-full bg-red-800" style={{ width: `${res.workload - 100}%` }}></div>)}
                                                            </div>
                                                            {isOverloaded && <p className="text-[10px] text-red-500 mt-1 font-medium">Beban melampaui batas wajar</p>}
                                                        </td>
                                                        <td className="p-4 text-right">
                                                            <div className="flex justify-end gap-2 items-center">
                                                                {/* TOMBOL DETAIL BARU */}
                                                                <button onClick={() => setViewingEmployee(res)} className="px-2 py-1.5 text-xs font-semibold bg-indigo-50 dark:bg-indigo-900/30 text-indigo-600 dark:text-indigo-400 hover:bg-indigo-100 dark:hover:bg-indigo-900/50 rounded-lg flex items-center gap-1.5 border border-indigo-200 dark:border-indigo-800/50 transition-colors">
                                                                    <Icon name="folder-open" size={14} /> Rincian
                                                                </button>
                                                                <div className="lg:opacity-0 lg:group-hover:opacity-100 transition-opacity flex gap-2">
                                                                    <button onClick={() => openModal('team', 'edit', res)} className="p-1.5 text-blue-600 dark:text-blue-400 hover:bg-blue-50 dark:hover:bg-blue-900/30 rounded transition-colors"><Icon name="edit" size={16} /></button>
                                                                    <button onClick={() => handleDelete('team', res.id)} className="p-1.5 text-red-600 dark:text-red-400 hover:bg-red-50 dark:hover:bg-red-900/30 rounded transition-colors"><Icon name="trash" size={16} /></button>
                                                                </div>
                                                            </div>
                                                        </td>
                                                    </tr>
                                                );
                                            })
                                        )}
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>
                );
            };

            // --- KOMPONEN TAB: INVENTARIS ---
            const renderInventory = () => {
                const filteredInv = inventory.filter(item => item.name?.toLowerCase().includes(searchInvTab.toLowerCase()) || item.type?.toLowerCase().includes(searchInvTab.toLowerCase()));

                return (
                    <div className="space-y-6 fade-in">
                        <div className="flex flex-col sm:flex-row justify-between gap-4">
                            <div className="relative flex-1">
                                <Icon name="search" className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={18} />
                                <input
                                    type="text"
                                    placeholder="Cari alat (nama atau kategori)..."
                                    value={searchInvTab}
                                    onChange={(e) => setSearchInvTab(e.target.value)}
                                    className="w-full pl-10 pr-4 py-2 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
                                />
                            </div>
                            <div className="flex gap-2">
                                {borrowCart.length > 0 && (
                                    <button
                                        onClick={() => setModalConfig({ isOpen: true, type: 'inventory-cart', mode: 'borrow-cart', data: null })}
                                        className="flex-shrink-0 flex items-center gap-2 px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white font-medium rounded-xl transition-colors shadow-sm text-sm"
                                    >
                                        <Icon name="shopping-cart" size={18} /> Pinjam {borrowCart.length} Alat
                                    </button>
                                )}
                            </div>
                        </div>

                        <div className="bg-white dark:bg-slate-800 rounded-2xl border border-slate-200 dark:border-slate-700 overflow-hidden shadow-sm">
                            <div className="overflow-x-auto">
                                <table className="w-full text-sm text-left">
                                    <thead className="text-xs text-slate-500 bg-slate-50 dark:bg-slate-800/50 uppercase border-b border-slate-200 dark:border-slate-700">
                                        <tr>
                                            <th className="px-6 py-4 w-12 text-center">
                                                <input
                                                    type="checkbox"
                                                    className="w-4 h-4 rounded border-gray-300 text-indigo-600 focus:ring-indigo-500 cursor-pointer"
                                                    checked={filteredInv.filter(i => i.status === 'Tersedia' && i.condition !== 'Rusak Berat').length > 0 && borrowCart.length === filteredInv.filter(i => i.status === 'Tersedia' && i.condition !== 'Rusak Berat').length}
                                                    onChange={(e) => {
                                                        if (e.target.checked) {
                                                            const availableIds = filteredInv.filter(i => i.status === 'Tersedia' && i.condition !== 'Rusak Berat').map(i => i.id);
                                                            setBorrowCart(availableIds);
                                                        } else {
                                                            setBorrowCart([]);
                                                        }
                                                    }}
                                                />
                                            </th>
                                            <th className="px-6 py-4 font-semibold">Nama Alat & ID</th>
                                            <th className="px-6 py-4 font-semibold">Kategori</th>
                                            <th className="px-6 py-4 font-semibold text-center">Kondisi</th>
                                            <th className="px-6 py-4 font-semibold text-center">Status</th>
                                            <th className="px-6 py-4 font-semibold text-right">Aksi</th>
                                        </tr>
                                    </thead>
                                    <tbody className="divide-y divide-slate-200 dark:divide-slate-700">
                                        {filteredInv.map((item) => {
                                            const isOverdue = (item.status === 'Dipinjam' || item.status === 'Menunggu Verifikasi' || item.status === 'Menunggu Verifikasi Pengembalian' || item.status === 'Menunggu Verifikasi Perpanjangan') && item.returnDate && new Date(item.returnDate).setHours(0, 0, 0, 0) < new Date().setHours(0, 0, 0, 0);
                                            return (
                                                <tr key={item.id} className={`transition-colors ${borrowCart.includes(item.id) ? 'bg-indigo-50 dark:bg-indigo-900/20' : 'hover:bg-slate-50 dark:hover:bg-slate-800/50'}`}>
                                                    <td className="px-6 py-4 text-center">
                                                        <input
                                                            type="checkbox"
                                                            className="w-4 h-4 rounded border-gray-300 text-indigo-600 focus:ring-indigo-500 cursor-pointer disabled:opacity-50 disabled:cursor-not-allowed"
                                                            disabled={item.status !== 'Tersedia' || item.condition === 'Rusak Berat'}
                                                            checked={borrowCart.includes(item.id)}
                                                            onChange={(e) => {
                                                                if (e.target.checked) {
                                                                    setBorrowCart([...borrowCart, item.id]);
                                                                } else {
                                                                    setBorrowCart(borrowCart.filter(id => id !== item.id));
                                                                }
                                                            }}
                                                        />
                                                    </td>
                                                    <td className="px-6 py-4">
                                                        <div className="font-semibold text-slate-900 dark:text-white">{item.name}</div>
                                                        <div className="text-xs text-slate-500 mt-1">{item.id}</div>
                                                    </td>
                                                    <td className="px-6 py-4 text-slate-600 dark:text-slate-300">{item.type}</td>
                                                    <td className="px-6 py-4 text-center">
                                                        <span className={`px-2.5 py-1 rounded-full text-xs font-medium ${item.condition === 'Baik' ? 'bg-green-50 text-green-700 dark:bg-green-900/30 dark:text-green-400' : 'bg-red-50 text-red-700 dark:bg-red-900/30 dark:text-red-400'}`}>
                                                            {item.condition}
                                                        </span>
                                                    </td>
                                                    <td className="px-6 py-4 text-center">
                                                        {item.status === 'Dipinjam' || item.status === 'Menunggu Verifikasi' || item.status === 'Menunggu Verifikasi Pengembalian' || item.status === 'Menunggu Verifikasi Perpanjangan' ? (
                                                            <div className="flex flex-col items-center">
                                                                <span className={`px-2.5 py-1 rounded-full text-xs font-medium ${isOverdue ? 'bg-red-50 text-red-700 dark:bg-red-900/30 dark:text-red-400 shadow-sm border border-red-100 dark:border-red-800' : (item.status === 'Menunggu Verifikasi' || item.status === 'Menunggu Verifikasi Pengembalian' || item.status === 'Menunggu Verifikasi Perpanjangan') ? 'bg-purple-50 text-purple-700 dark:bg-purple-900/30 dark:text-purple-400' : 'bg-amber-50 text-amber-700 dark:bg-amber-900/30 dark:text-amber-400'}`}>
                                                                    {isOverdue ? 'Masa pinjam habis' : item.status}
                                                                </span>
                                                                <div className={`text-[11px] mt-2 text-center whitespace-nowrap ${isOverdue ? 'text-red-600 dark:text-red-400 font-semibold' : 'text-slate-500'}`}>
                                                                    Oleh: <strong>{item.borrower}</strong><br />
                                                                    {item.borrowDate ? `${formatDateIndo(item.borrowDate)} s/d ` : ''}{item.returnDate ? formatDateIndo(item.returnDate) : '-'}
                                                                    {item.status === 'Menunggu Verifikasi Perpanjangan' && item.newReturnDate && (
                                                                        <div className="text-amber-600 font-semibold mt-1">Request: {formatDateIndo(item.newReturnDate)}</div>
                                                                    )}
                                                                </div>
                                                            </div>
                                                        ) : (
                                                            <div className="flex flex-col items-center">
                                                                <span className="px-2.5 py-1 rounded-full text-xs font-medium bg-emerald-50 text-emerald-700 dark:bg-emerald-900/30 dark:text-emerald-400">
                                                                    Tersedia
                                                                </span>
                                                            </div>
                                                        )}
                                                    </td>
                                                    <td className="px-6 py-4 text-right">
                                                        <div className="flex items-center justify-end gap-2">
                                                            {item.status === 'Tersedia' ? (
                                                                <button
                                                                    onClick={() => setModalConfig({ isOpen: true, type: 'inventory-borrow', mode: 'borrow', data: item })}
                                                                    disabled={item.condition === 'Rusak Berat'}
                                                                    title={item.condition === 'Rusak Berat' ? "Alat rusak, tidak dapat dipinjam." : "Pinjam"}
                                                                    className={`px-3 py-1.5 text-xs font-medium rounded-lg transition-colors border ${item.condition === 'Rusak Berat' ? 'bg-slate-100 text-slate-400 border-slate-200 cursor-not-allowed dark:bg-slate-800 dark:text-slate-500 dark:border-slate-700' : 'bg-indigo-50 text-indigo-700 hover:bg-indigo-100 dark:bg-indigo-900/30 dark:text-indigo-400 border-indigo-200 dark:border-indigo-800'}`}
                                                                >
                                                                    Pinjam
                                                                </button>
                                                            ) : item.status === 'Menunggu Verifikasi' || item.status === 'Menunggu Verifikasi Pengembalian' || item.status === 'Menunggu Verifikasi Perpanjangan' ? (
                                                                <span className="text-xs text-purple-600 font-medium px-2 py-1 bg-purple-50 rounded-lg">Menunggu Logistik</span>
                                                            ) : (
                                                                <>
                                                                    <button
                                                                        onClick={() => setModalConfig({ isOpen: true, type: 'inventory-extend', mode: 'extend', data: item })}
                                                                        className="px-3 py-1.5 text-xs font-medium bg-amber-50 text-amber-700 hover:bg-amber-100 dark:bg-amber-900/30 dark:text-amber-400 rounded-lg transition-colors border border-amber-200 dark:border-amber-800"
                                                                    >
                                                                        Perpanjang
                                                                    </button>
                                                                    <button
                                                                        onClick={() => setModalConfig({ isOpen: true, type: 'inventory-return', mode: 'return', data: item })}
                                                                        className="px-3 py-1.5 text-xs font-medium bg-white text-slate-700 hover:bg-slate-50 border border-slate-300 dark:bg-slate-800 dark:border-slate-600 dark:text-slate-300 rounded-lg transition-colors"
                                                                    >
                                                                        Kembalikan
                                                                    </button>
                                                                </>
                                                            )}
                                                        </div>
                                                    </td>
                                                </tr>
                                            );
                                        })}
                                        {filteredInv.length === 0 && (
                                            <tr>
                                                <td colSpan="5" className="px-6 py-12 text-center">
                                                    <div className="flex flex-col items-center justify-center text-slate-400 dark:text-slate-500">
                                                        <Icon name="box" size={48} className="mb-4 opacity-50" />
                                                        <p className="text-sm">Tidak ada data alat/inventaris ditemukan.</p>
                                                    </div>
                                                </td>
                                            </tr>
                                        )}
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>
                );
            };

            // --- KOMPONEN TAB: TIMESHEET ---
            const renderTimesheet = () => (
                <div className="flex flex-col items-center justify-center h-[60vh] text-center fade-in">
                    <div className="w-20 h-20 bg-indigo-50 dark:bg-indigo-900/30 text-indigo-600 dark:text-indigo-400 rounded-full flex items-center justify-center mb-4">
                        <Icon name="calendar-clock" size={40} />
                    </div>
                    <h2 className="text-2xl font-bold text-slate-800 dark:text-slate-100 mb-2">Modul Timesheet</h2>
                    <p className="text-slate-500 dark:text-slate-400 max-w-md">Modul pengisian log harian karyawan sedang dalam tahap pengembangan dan akan segera dirilis pada pembaruan sistem berikutnya.</p>
                </div>
            );

            const ExpertModalForm = () => {
                if (!modalConfig.isOpen || modalConfig.type !== 'expert') return null;
                const isEdit = modalConfig.mode === 'edit';

                const [formData, setFormData] = useState(() => {
                    if (isEdit && modalConfig.data) return { ...modalConfig.data };
                    return { id: '', name: '', phone: '', status: 'Tersedia', jenjang: '', bidangIlmu: '', perusahaan: '', linkedResourceName: '' };
                });

                const handleSubmit = (e) => {
                    e.preventDefault();
                    handleExpertAction(isEdit ? 'edit' : 'add', formData);
                };

                return (
                    <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[100] flex items-center justify-center fade-in p-4">
                        <div className="glass-card rounded-3xl shadow-2xl w-full max-w-lg flex flex-col max-h-[90vh] overflow-hidden transform scale-in border border-slate-200 dark:border-slate-800">
                            <div className="flex justify-between items-center p-5 border-b border-slate-100 dark:border-slate-700/50 bg-slate-50/50 dark:bg-slate-900/50">
                                <h3 className="text-xl font-bold text-slate-800 dark:text-slate-100">{isEdit ? 'Edit Data Tenaga Ahli' : 'Tambah Tenaga Ahli'}</h3>
                                <button onClick={() => setModalConfig({ isOpen: false, type: null })} type="button" className="text-slate-400 hover:bg-slate-200 hover:text-slate-600 dark:hover:bg-slate-700 dark:hover:text-slate-300 p-2 rounded-xl transition-colors"><Icon name="x" size={20} /></button>
                            </div>
                            <div className="p-6 overflow-y-auto custom-scrollbar flex-1">
                                <form id="expertForm" onSubmit={handleSubmit} className="space-y-4">
                                    <div>
                                        <label className="block text-xs font-bold text-slate-700 dark:text-slate-300 mb-1">Nama Lengkap</label>
                                        <input type="text" required value={formData.name} onChange={e => setFormData({ ...formData, name: e.target.value })} className="w-full p-3 rounded-xl border border-slate-200 dark:border-slate-700/50 focus:border-indigo-500 focus:ring-2 focus:ring-indigo-500/20 outline-none text-sm bg-slate-50/50 dark:bg-slate-900/50 text-slate-800 dark:text-slate-200 shadow-inner transition-all" />
                                    </div>
                                    <div>
                                        <label className="block text-xs font-bold text-slate-700 dark:text-slate-300 mb-1">No. HP / WhatsApp</label>
                                        <input type="text" value={formData.phone} onChange={e => setFormData({ ...formData, phone: e.target.value })} className="w-full p-3 rounded-xl border border-slate-200 dark:border-slate-700/50 focus:border-indigo-500 focus:ring-2 focus:ring-indigo-500/20 outline-none text-sm bg-slate-50/50 dark:bg-slate-900/50 text-slate-800 dark:text-slate-200 shadow-inner transition-all" />
                                    </div>
                                    <div>
                                        <label className="block text-xs font-bold text-slate-700 dark:text-slate-300 mb-1">Jenjang & Bidang Ilmu</label>
                                        <input type="text" placeholder="Cth: SMA, S1 Teknik Sipil Thn 1997" value={formData.bidangIlmu || ''} onChange={e => setFormData({ ...formData, bidangIlmu: e.target.value })} className="w-full p-3 rounded-xl border border-slate-200 dark:border-slate-700/50 focus:border-indigo-500 focus:ring-2 focus:ring-indigo-500/20 outline-none text-sm bg-slate-50/50 dark:bg-slate-900/50 text-slate-800 dark:text-slate-200 shadow-inner transition-all" />
                                    </div>
                                    <div>
                                        <label className="block text-xs font-bold text-slate-700 dark:text-slate-300 mb-1">Perusahaan / Instansi</label>
                                        <select value={formData.perusahaan || ''} onChange={e => setFormData({ ...formData, perusahaan: e.target.value })} className="w-full p-3 rounded-xl border border-slate-200 dark:border-slate-700/50 focus:border-indigo-500 focus:ring-2 focus:ring-indigo-500/20 outline-none text-sm bg-slate-50/50 dark:bg-slate-900/50 text-slate-800 dark:text-slate-200 shadow-inner transition-all">
                                            <option value="">-- Pilih Perusahaan --</option>
                                            <option value="PT. Gaharu Sempana">PT. Gaharu Sempana</option>
                                            <option value="PT. Kencana Adhi Karma">PT. Kencana Adhi Karma</option>
                                            <option value="CV. Cipta Asri Disain">CV. Cipta Asri Disain</option>
                                            <option value="CV. Tataring Bali">CV. Tataring Bali</option>
                                            <option value="Freelance">Freelance</option>
                                        </select>
                                    </div>
                                    <div>
                                        <label className="block text-xs font-bold text-slate-700 dark:text-slate-300 mb-1 flex items-center gap-1.5"><Icon name="link" size={14} className="text-indigo-500"/> Hubungkan dgn Personil Internal (Opsional)</label>
                                        <input type="text" list="internal-resources-list" placeholder="-- Ketik untuk mencari personil atau biarkan kosong --" value={formData.linkedResourceName || ''} onChange={e => setFormData({ ...formData, linkedResourceName: e.target.value })} className="w-full p-3 rounded-xl border border-indigo-200 dark:border-indigo-700/50 focus:border-indigo-500 focus:ring-2 focus:ring-indigo-500/20 outline-none text-sm bg-indigo-50/30 dark:bg-indigo-900/10 text-slate-800 dark:text-slate-200 shadow-inner transition-all" />
                                        <datalist id="internal-resources-list">
                                            {resources.map(r => (
                                                <option key={r.id || r.name} value={r.name}>{r.name} ({r.role})</option>
                                            ))}
                                        </datalist>
                                        <p className="text-[10px] text-slate-500 mt-1">Pilih ini jika nama Tenaga Ahli di kontrak berbeda dengan nama di menu Alokasi Tim agar tetap terbaca & tersinkron pada beban kerja.</p>
                                    </div>
                                </form>
                            </div>
                            <div className="p-5 border-t border-slate-100 dark:border-slate-700/50 flex justify-end gap-3 bg-slate-50 dark:bg-slate-900/50">
                                <button onClick={() => setModalConfig({ isOpen: false, type: null })} type="button" className="px-4 py-2 text-sm font-medium text-slate-600 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-700 rounded-xl">Batal</button>
                                <button form="expertForm" type="submit" disabled={loading} className="px-6 py-2 bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-medium rounded-xl disabled:opacity-70 flex items-center gap-2">
                                    {loading ? <Icon name="refresh-ccw" className="animate-spin" size={16} /> : <Icon name="save" size={16} />} Simpan
                                </button>
                            </div>
                        </div>
                    </div>
                );
            };

            const ExpertCertModalForm = () => {
                if (!modalConfig.isOpen || modalConfig.type !== 'expert_cert') return null;
                const { expertId, certIndex, cert } = modalConfig.data;
                const expert = experts.find(e => e.id === expertId);
                const isEdit = modalConfig.mode === 'edit';

                const [formData, setFormData] = useState(() => {
                    if (isEdit && cert) return { ...cert };
                    return { certName: '', certLevel: 'Ahli Muda', issuedDate: '', expiredDate: '' };
                });

                const handleSubmit = (e) => {
                    e.preventDefault();
                    if (!expert) return;
                    let newCerts = [...(expert.certificates || [])];
                    if (isEdit) {
                        newCerts[certIndex] = formData;
                    } else {
                        newCerts.push(formData);
                    }
                    handleExpertAction('update_certificates', { ...expert, certificates: newCerts });
                };

                return (
                    <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[100] flex items-center justify-center fade-in p-4">
                        <div className="glass-card rounded-3xl shadow-2xl w-full max-w-md flex flex-col overflow-hidden transform scale-in border border-slate-200 dark:border-slate-800">
                            <div className="flex justify-between items-center p-5 border-b border-slate-100 dark:border-slate-700/50 bg-slate-50/50 dark:bg-slate-900/50">
                                <h3 className="text-xl font-bold text-slate-800 dark:text-slate-100">{isEdit ? 'Edit Sertifikat' : 'Tambah Sertifikat'}</h3>
                                <button onClick={() => setModalConfig({ isOpen: false, type: null })} type="button" className="text-slate-400 hover:bg-slate-200 hover:text-slate-600 dark:hover:bg-slate-700 dark:hover:text-slate-300 p-2 rounded-xl transition-colors"><Icon name="x" size={20} /></button>
                            </div>
                            <div className="p-6">
                                <form id="expertCertForm" onSubmit={handleSubmit} className="space-y-4">
                                    <div>
                                        <div className="flex justify-between items-center mb-1">
                                            <label className="block text-xs font-bold text-slate-700 dark:text-slate-300">Nama Sertifikat</label>
                                            <button type="button" onClick={() => setShowCertManager(true)} className="text-xs text-indigo-600 dark:text-indigo-400 hover:text-indigo-700 font-semibold flex items-center gap-1">
                                                <Icon name="settings" size={12} /> Kelola Daftar
                                            </button>
                                        </div>
                                        <input type="text" required list="cert-options" value={formData.certName} onChange={e => setFormData({ ...formData, certName: e.target.value })} placeholder="Cth: SKA Ahli Teknik Bangunan Gedung" className="w-full p-3 rounded-xl border border-slate-200 dark:border-slate-700/50 focus:border-indigo-500 focus:ring-2 focus:ring-indigo-500/20 outline-none text-sm bg-slate-50/50 dark:bg-slate-900/50 text-slate-800 dark:text-slate-200 shadow-inner transition-all" />
                                        <datalist id="cert-options">
                                            {certList.map((cert, idx) => <option key={idx} value={cert} />)}
                                        </datalist>
                                    </div>
                                    <div>
                                        <label className="block text-xs font-bold text-slate-700 dark:text-slate-300 mb-1">Tingkat / Kualifikasi</label>
                                        <select value={formData.certLevel} onChange={e => setFormData({ ...formData, certLevel: e.target.value })} className="w-full p-3 rounded-xl border border-slate-200 dark:border-slate-700/50 focus:border-indigo-500 focus:ring-2 focus:ring-indigo-500/20 outline-none text-sm bg-slate-50/50 dark:bg-slate-900/50 text-slate-800 dark:text-slate-200 shadow-inner transition-all">
                                            <option value="Ahli Muda">Ahli Muda</option>
                                            <option value="Ahli Madya">Ahli Madya</option>
                                            <option value="Ahli Utama">Ahli Utama</option>
                                            <option value="Lainnya">Lainnya</option>
                                        </select>
                                    </div>
                                    <div className="grid grid-cols-2 gap-4">
                                        <div>
                                            <label className="block text-xs font-bold text-slate-700 dark:text-slate-300 mb-1">Tgl Terbit</label>
                                            <input type="date" required value={formData.issuedDate} onChange={e => setFormData({ ...formData, issuedDate: e.target.value })} className="w-full p-3 rounded-xl border border-slate-200 dark:border-slate-700/50 focus:border-indigo-500 focus:ring-2 focus:ring-indigo-500/20 outline-none text-sm bg-slate-50/50 dark:bg-slate-900/50 text-slate-800 dark:text-slate-200 shadow-inner transition-all" />
                                        </div>
                                        <div>
                                            <label className="block text-xs font-bold text-slate-700 dark:text-slate-300 mb-1">Tgl Berakhir</label>
                                            <input type="date" required value={formData.expiredDate} onChange={e => setFormData({ ...formData, expiredDate: e.target.value })} className="w-full p-3 rounded-xl border border-slate-200 dark:border-slate-700/50 focus:border-indigo-500 focus:ring-2 focus:ring-indigo-500/20 outline-none text-sm bg-slate-50/50 dark:bg-slate-900/50 text-slate-800 dark:text-slate-200 shadow-inner transition-all" />
                                        </div>
                                    </div>
                                </form>
                            </div>
                            <div className="p-5 border-t border-slate-100 dark:border-slate-700/50 flex justify-end gap-3 bg-slate-50 dark:bg-slate-900/50">
                                <button onClick={() => setModalConfig({ isOpen: false, type: null })} type="button" className="px-4 py-2 text-sm font-medium text-slate-600 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-700 rounded-xl">Batal</button>
                                <button form="expertCertForm" type="submit" disabled={loading} className="px-6 py-2 bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-medium rounded-xl disabled:opacity-70 flex items-center gap-2">
                                    {loading ? <Icon name="refresh-ccw" className="animate-spin" size={16} /> : <Icon name="save" size={16} />} Simpan
                                </button>
                            </div>
                        </div>
                    </div>
                );
            };

            const ExpertTenderModalForm = () => {
                if (!modalConfig.isOpen || modalConfig.type !== 'expert_tender') return null;
                const { expertId, tenderIndex, tender } = modalConfig.data;
                const expert = experts.find(e => e.id === expertId);
                const isEdit = modalConfig.mode === 'edit';

                const [formData, setFormData] = useState(() => {
                    if (isEdit && tender) return { ...tender };
                    return { lpseName: '', position: 'Team Leader', status: 'Aktif' };
                });

                const handleSubmit = (e) => {
                    e.preventDefault();
                    if (!expert) return;

                    // Validation rule
                    if (formData.status === 'Aktif' || formData.status === 'Menunggu Pengumuman') {
                        const existingActive = (expert.tenders || []).find((t, idx) =>
                            (!isEdit || idx !== tenderIndex) &&
                            t.lpseName.trim().toLowerCase() === formData.lpseName.trim().toLowerCase() &&
                            (t.status === 'Aktif' || t.status === 'Menunggu Pengumuman')
                        );

                        if (existingActive) {
                            alert(`Gagal menyimpan! Tenaga Ahli ini sudah memiliki tender aktif/menunggu di ${existingActive.lpseName}. Sesuai aturan, tidak boleh dipasang pada lebih dari 1 tender aktif di LPSE yang sama.`);
                            return;
                        }
                    }

                    let newTenders = [...(expert.tenders || [])];
                    if (isEdit) {
                        newTenders[tenderIndex] = formData;
                    } else {
                        newTenders.push(formData);
                    }
                    handleExpertAction('update_tenders', { ...expert, tenders: newTenders });
                };

                return (
                    <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[100] flex items-center justify-center fade-in p-4">
                        <div className="glass-card rounded-3xl shadow-2xl w-full max-w-md flex flex-col overflow-hidden transform scale-in border border-slate-200 dark:border-slate-800">
                            <div className="flex justify-between items-center p-5 border-b border-slate-100 dark:border-slate-700/50 bg-slate-50/50 dark:bg-slate-900/50">
                                <h3 className="text-xl font-bold text-slate-800 dark:text-slate-100">{isEdit ? 'Edit Riwayat Tender' : 'Plotting Tender LPSE'}</h3>
                                <button onClick={() => setModalConfig({ isOpen: false, type: null })} type="button" className="text-slate-400 hover:bg-slate-200 hover:text-slate-600 dark:hover:bg-slate-700 dark:hover:text-slate-300 p-2 rounded-xl transition-colors"><Icon name="x" size={20} /></button>
                            </div>
                            <div className="p-6">
                                <form id="expertTenderForm" onSubmit={handleSubmit} className="space-y-4">
                                    <div>
                                        <label className="block text-xs font-bold text-slate-700 dark:text-slate-300 mb-1">Nama / Instansi LPSE</label>
                                        <div className="flex gap-2 items-center">
                                            <input type="text" list="lpse-options" required value={formData.lpseName} onChange={e => setFormData({ ...formData, lpseName: e.target.value })} placeholder="Cth: LPSE Kementerian PUPR" className="w-full p-3 rounded-xl border border-slate-200 dark:border-slate-700/50 focus:border-indigo-500 focus:ring-2 focus:ring-indigo-500/20 outline-none text-sm bg-slate-50/50 dark:bg-slate-900/50 text-slate-800 dark:text-slate-200 shadow-inner transition-all" />
                                            <datalist id="lpse-options">
                                                {lpseList.map((lpse, idx) => <option key={idx} value={lpse} />)}
                                            </datalist>
                                            <button type="button" onClick={() => setShowLpseManager(true)} className="p-2.5 bg-slate-100 dark:bg-slate-800 hover:bg-slate-200 dark:hover:bg-slate-700 text-slate-500 rounded-xl transition-colors shrink-0" title="Kelola Daftar LPSE">
                                                <Icon name="settings" size={18} />
                                            </button>
                                        </div>
                                    </div>
                                    <div>
                                        <label className="block text-xs font-bold text-slate-700 dark:text-slate-300 mb-1">Posisi / Jabatan yang Ditawarkan</label>
                                        <input type="text" required value={formData.position} onChange={e => setFormData({ ...formData, position: e.target.value })} placeholder="Cth: Team Leader / Ahli Struktur" className="w-full p-3 rounded-xl border border-slate-200 dark:border-slate-700/50 focus:border-indigo-500 focus:ring-2 focus:ring-indigo-500/20 outline-none text-sm bg-slate-50/50 dark:bg-slate-900/50 text-slate-800 dark:text-slate-200 shadow-inner transition-all" />
                                    </div>
                                    <div>
                                        <label className="block text-xs font-bold text-slate-700 dark:text-slate-300 mb-1">Status Tender</label>
                                        <select value={formData.status} onChange={e => setFormData({ ...formData, status: e.target.value })} className="w-full p-3 rounded-xl border border-slate-200 dark:border-slate-700/50 focus:border-indigo-500 focus:ring-2 focus:ring-indigo-500/20 outline-none text-sm bg-slate-50/50 dark:bg-slate-900/50 text-slate-800 dark:text-slate-200 shadow-inner transition-all">
                                            <option value="Menunggu Pengumuman">Menunggu Pengumuman</option>
                                            <option value="Aktif">Aktif (Sedang Proses)</option>
                                            <option value="Menang">Menang</option>
                                            <option value="Kalah">Kalah / Gugur</option>
                                            <option value="Selesai">Selesai / Riwayat</option>
                                        </select>
                                    </div>
                                </form>
                            </div>
                            <div className="p-5 border-t border-slate-100 dark:border-slate-700/50 flex justify-end gap-3 bg-slate-50 dark:bg-slate-900/50">
                                <button onClick={() => setModalConfig({ isOpen: false, type: null })} type="button" className="px-4 py-2 text-sm font-medium text-slate-600 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-700 rounded-xl">Batal</button>
                                <button form="expertTenderForm" type="submit" disabled={loading} className="px-6 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-medium rounded-xl disabled:opacity-70 flex items-center gap-2">
                                    {loading ? <Icon name="refresh-ccw" className="animate-spin" size={16} /> : <Icon name="save" size={16} />} Simpan
                                </button>
                            </div>
                        </div>
                    </div>
                );
            };

            // --- IMPORT EXCEL MODAL ---
            const ImportExcelModal = () => {
                if (!modalConfig.isOpen || modalConfig.type !== 'import_expert') return null;
                
                return (
                    <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[100] flex items-center justify-center fade-in p-4">
                        <div className="glass-card rounded-3xl shadow-2xl w-full max-w-sm flex flex-col overflow-hidden transform scale-in border border-slate-200 dark:border-slate-800 p-6">
                            <div className="flex justify-between items-center mb-6">
                                <h3 className="text-xl font-bold text-slate-800 dark:text-slate-100">Import Data Excel</h3>
                                <button onClick={() => setModalConfig({ isOpen: false, type: null })} className="text-slate-400 hover:bg-slate-200 hover:text-slate-600 dark:hover:bg-slate-700 dark:hover:text-slate-300 p-2 rounded-xl transition-colors"><Icon name="x" size={20} /></button>
                            </div>
                            <div className="space-y-4">
                                <p className="text-sm text-slate-500 dark:text-slate-400">Silakan download template format Excel yang didukung sistem terlebih dahulu untuk menghindari kesalahan input.</p>
                                
                                <button onClick={() => {
                                    const ws_data = [
                                        ["Nama Personil", "No HP", "Status", "Jenjang dan Bidang Ilmu", "Perusahaan", "Sertifikat Keahlian", "Jenjang", "Masa Berlaku"],
                                        ["Contoh: Budi Santoso", "08123456789", "Tersedia", "S1 Teknik Sipil", "PT. Maju Jaya", "Ahli Teknik Bangunan Gedung", "Madya", "2027-12-31"]
                                    ];
                                    const ws = XLSX.utils.aoa_to_sheet(ws_data);
                                    ws['!cols'] = [{wch: 25}, {wch: 15}, {wch: 15}, {wch: 25}, {wch: 20}, {wch: 30}, {wch: 15}, {wch: 15}];
                                    const wb = XLSX.utils.book_new();
                                    XLSX.utils.book_append_sheet(wb, ws, "Template");
                                    XLSX.writeFile(wb, "Template_Import_Tenaga_Ahli.xlsx");
                                }} className="w-full px-4 py-3 bg-slate-100 hover:bg-slate-200 dark:bg-slate-800 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-200 rounded-xl font-semibold flex items-center justify-center gap-2 text-sm transition-colors border border-slate-200 dark:border-slate-700 shadow-sm">
                                    <Icon name="download" size={18} /> Download Template Excel
                                </button>

                                <div className="relative flex py-2 items-center">
                                    <div className="flex-grow border-t border-slate-200 dark:border-slate-700"></div>
                                    <span className="flex-shrink-0 mx-4 text-slate-400 text-xs font-semibold">ATAU</span>
                                    <div className="flex-grow border-t border-slate-200 dark:border-slate-700"></div>
                                </div>

                                <label className="w-full cursor-pointer px-4 py-3 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl font-semibold flex items-center justify-center gap-2 text-sm transition-colors shadow-md">
                                    <Icon name="upload" size={18} /> Pilih File & Import
                                    <input type="file" accept=".xlsx, .xls" onChange={(e) => {
                                        setModalConfig({ isOpen: false, type: null });
                                        handleImportExcel(e);
                                    }} className="hidden" />
                                </label>
                            </div>
                        </div>
                    </div>
                );
            };

            // --- ASSIGNMENT MODAL FORM ---
            const AssignmentModalForm = () => {
                if (!modalConfig.isOpen || modalConfig.type !== 'assignment') return null;

                const isEdit = modalConfig.mode === 'edit';

                const [formData, setFormData] = useState(() => {
                    if (isEdit && modalConfig.data) {
                        return { ...modalConfig.data, experts: modalConfig.data.experts || [] };
                    }
                    return { jobName: '', projectType: 'Pengawasan', contractType: 'Waktu Penugasan', tenderType: 'Tender', lpseName: '', startDate: '', duration: '', experts: [] };
                });

                // Auto hitung end date jika start date & durasi diisi
                const [endDate, setEndDate] = useState('');
                useEffect(() => {
                    if (formData.startDate && formData.duration) {
                        const start = new Date(formData.startDate);
                        const days = parseInt(formData.duration, 10);
                        if (!isNaN(days)) {
                            const end = new Date(start);
                            end.setDate(start.getDate() + days);
                            setEndDate(end.toISOString().split('T')[0]);
                        } else {
                            setEndDate('');
                        }
                    } else {
                        setEndDate('');
                    }
                }, [formData.startDate, formData.duration]);

                const handleSubmit = (e) => {
                    e.preventDefault();
                    if (!formData.jobName || !formData.lpseName) {
                        alert("Harap lengkapi informasi pekerjaan!");
                        return;
                    }
                    if (!formData.experts || formData.experts.length === 0) {
                        alert("Harap plot minimal 1 tenaga ahli ke pekerjaan ini!");
                        return;
                    }

                    for (let i = 0; i < formData.experts.length; i++) {
                        const exp = formData.experts[i];
                        if (!exp.expertId || !exp.manMonth) {
                            setAlertModal({ isOpen: true, title: 'Validasi Form', message: `Harap lengkapi data personil (Tenaga Ahli dan Man Month) di baris ke-${i + 1}` });
                            return;
                        }

                        // Cek aturan perundang-undangan LPSE
                        const expertObj = experts.find(e => e.id === exp.expertId);
                        const expertName = expertObj ? expertObj.name : 'Tenaga Ahli';

                        const existingAssignmentsInSameLpse = assignments.filter(asg => {
                            if (isEdit && asg.id === formData.id) return false;
                            if ((asg.lpseName || '').trim().toLowerCase() !== (formData.lpseName || '').trim().toLowerCase()) return false;

                            // Abaikan pekerjaan yang sudah expired / lewat dari hari ini
                            if (asg.endDate) {
                                const today = new Date();
                                today.setHours(0, 0, 0, 0);
                                const asgEnd = new Date(asg.endDate);
                                asgEnd.setHours(0, 0, 0, 0);
                                if (asgEnd < today) {
                                    return false; // Pekerjaan sudah selesai dan dianggap hilang, jadi tidak masuk kuota
                                }
                            }

                            // Cek overlapping tanggal (jika tanggal akhir pekerjaan lama sudah lewat dari tanggal mulai pekerjaan baru, maka aman dan tidak masuk perhitungan kuota)
                            if (formData.startDate && endDate && asg.startDate && asg.endDate) {
                                const newStart = new Date(formData.startDate);
                                const newEnd = new Date(endDate);
                                const asgStart = new Date(asg.startDate);
                                const asgEnd = new Date(asg.endDate);

                                // Reset time to 00:00:00 to prevent edge case issues
                                newStart.setHours(0, 0, 0, 0);
                                newEnd.setHours(0, 0, 0, 0);
                                asgStart.setHours(0, 0, 0, 0);
                                asgEnd.setHours(0, 0, 0, 0);

                                if (newStart > asgEnd || newEnd < asgStart) {
                                    return false; // Tidak tumpang tindih, boleh diplot
                                }
                            }

                            return (asg.experts || []).some(e => e.expertId === exp.expertId);
                        });

                        const countWaktuPenugasan = existingAssignmentsInSameLpse.filter(asg => (asg.contractType || 'Waktu Penugasan') === 'Waktu Penugasan').length;
                        const countLumsum = existingAssignmentsInSameLpse.filter(asg => asg.contractType === 'Lumsum').length;

                        if (countWaktuPenugasan >= 1) {
                            setAlertModal({ isOpen: true, title: 'Kapasitas Penuh', message: `Gagal menyimpan! ${expertName} sudah memiliki 1 pekerjaan Waktu Penugasan di ${formData.lpseName}.\n\nSesuai aturan, Tenaga Ahli tidak dapat di-plot di pekerjaan baru (Waktu Penugasan / Lumsum) pada LPSE yang sama.` });
                            return;
                        }
                        if (countLumsum >= 3) {
                            setAlertModal({ isOpen: true, title: 'Kapasitas Penuh', message: `Gagal menyimpan! ${expertName} sudah memiliki 3 pekerjaan Lumsum di ${formData.lpseName}.\n\nSesuai aturan, Tenaga Ahli telah mencapai batas maksimal 3 pekerjaan dan tidak dapat di-plot di pekerjaan baru pada LPSE yang sama.` });
                            return;
                        }

                        if ((formData.contractType || 'Waktu Penugasan') === 'Waktu Penugasan') {
                            if (countLumsum > 0) {
                                setAlertModal({ isOpen: true, title: 'Melanggar Batas Aturan', message: `Gagal menyimpan! ${expertName} sudah memiliki pekerjaan Lumsum di ${formData.lpseName}.\n\nTenaga ahli tidak dapat ditambahkan ke pekerjaan Waktu Penugasan karena tidak dapat digabungkan dengan kontrak Lumsum yang sudah berjalan.` });
                                return;
                            }
                        }
                    }
                    const cleanedExperts = formData.experts.map(exp => {
                        const { expertSearchTemp, ...rest } = exp;
                        return rest;
                    });

                    const payload = { ...formData, experts: cleanedExperts, endDate };
                    handleAssignmentAction(isEdit ? 'edit' : 'add', payload);
                };

                const handleAddExpertRow = () => {
                    setFormData(prev => ({
                        ...prev,
                        experts: [...(prev.experts || []), { expertId: '', role: '', certificateName: '', additionalCertificates: [], manMonth: '' }]
                    }));
                };

                const handleRemoveExpertRow = (index) => {
                    setFormData(prev => {
                        const newExperts = [...prev.experts];
                        newExperts.splice(index, 1);
                        return { ...prev, experts: newExperts };
                    });
                };

                const handleExpertRowChange = (index, field, value) => {
                    setFormData(prev => {
                        const newExperts = [...prev.experts];
                        newExperts[index] = { ...newExperts[index], [field]: value };

                        // Reset certificate when expert changes
                        if (field === 'expertId') {
                            newExperts[index].certificateName = '';
                            newExperts[index].additionalCertificates = [];
                        }

                        return { ...prev, experts: newExperts };
                    });
                };

                return (
                    <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[100] flex items-center justify-center fade-in p-4">
                        <div className="glass-card rounded-3xl shadow-2xl w-full max-w-3xl flex flex-col max-h-[90vh] overflow-hidden transform scale-in border border-slate-200 dark:border-slate-800">
                            <div className="flex justify-between items-center p-5 border-b border-slate-100 dark:border-slate-700/50 bg-slate-50/50 dark:bg-slate-900/50">
                                <h3 className="text-xl font-bold text-slate-800 dark:text-slate-100">{isEdit ? 'Edit Penugasan' : 'Tambah Penugasan'}</h3>
                                <button onClick={() => setModalConfig({ isOpen: false, type: null })} type="button" className="text-slate-400 hover:bg-slate-200 hover:text-slate-600 dark:hover:bg-slate-700 dark:hover:text-slate-300 p-2 rounded-xl transition-colors"><Icon name="x" size={20} /></button>
                            </div>
                            <div className="p-6 overflow-y-auto custom-scrollbar flex-1">
                                <form id="assignmentForm" onSubmit={handleSubmit} className="space-y-6">

                                    {/* SEKSI INFORMASI PEKERJAAN */}
                                    <div className="space-y-4">
                                        <h4 className="font-bold text-slate-700 dark:text-slate-200 border-b border-slate-200 dark:border-slate-700 pb-2">Informasi Pekerjaan</h4>
                                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                            <div>
                                                <label className="block text-xs font-bold text-slate-700 dark:text-slate-300 mb-1">Nama Pekerjaan</label>
                                                <input type="text" required value={formData.jobName} onChange={e => setFormData({ ...formData, jobName: e.target.value })} className="w-full p-3 rounded-xl border border-slate-200 dark:border-slate-700/50 focus:border-indigo-500 focus:ring-2 focus:ring-indigo-500/20 outline-none text-sm bg-slate-50/50 dark:bg-slate-900/50 text-slate-800 dark:text-slate-200 shadow-inner transition-all" />
                                            </div>
                                            <div>
                                                <label className="block text-xs font-bold text-slate-700 dark:text-slate-300 mb-1">Klien / Owner Proyek</label>
                                                <input type="text" value={formData.clientName || ''} onChange={e => setFormData({ ...formData, clientName: e.target.value })} placeholder={formData.lpseName || 'Cth: Dinas PUPR Provinsi Bali'} className="w-full p-3 rounded-xl border border-slate-200 dark:border-slate-700/50 focus:border-indigo-500 focus:ring-2 focus:ring-indigo-500/20 outline-none text-sm bg-slate-50/50 dark:bg-slate-900/50 text-slate-800 dark:text-slate-200 shadow-inner transition-all" />
                                                <p className="text-[10px] text-slate-400 mt-1">Jika kosong, akan menggunakan nama LPSE. Tampil di List Proyek.</p>
                                            </div>
                                        </div>
                                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                            <div className="sm:col-span-2">
                                                <label className="block text-xs font-bold text-slate-700 dark:text-slate-300 mb-1">LPSE / Instansi</label>
                                                <div className="flex gap-2 items-center">
                                                    <input type="text" list="assignment-lpse-options" required value={formData.lpseName} onChange={e => setFormData({ ...formData, lpseName: e.target.value })} placeholder="Cth: LPSE Provinsi Bali" className="w-full p-3 rounded-xl border border-slate-200 dark:border-slate-700/50 focus:border-indigo-500 focus:ring-2 focus:ring-indigo-500/20 outline-none text-sm bg-slate-50/50 dark:bg-slate-900/50 text-slate-800 dark:text-slate-200 shadow-inner transition-all" />
                                                    <datalist id="assignment-lpse-options">
                                                        {lpseList.map((lpse, idx) => <option key={idx} value={lpse} />)}
                                                    </datalist>
                                                    <button type="button" onClick={() => setShowLpseManager(true)} className="p-2.5 bg-slate-100 dark:bg-slate-800 hover:bg-slate-200 dark:hover:bg-slate-700 text-slate-500 rounded-xl transition-colors shrink-0" title="Kelola Daftar LPSE">
                                                        <Icon name="settings" size={18} />
                                                    </button>
                                                </div>
                                            </div>
                                        </div>
                                        {/* Banner info sinkronisasi otomatis */}
                                        {(formData.projectType === 'Pengawasan' || formData.projectType === 'Manajemen Konstruksi') && (
                                            <div className="flex items-start gap-2.5 px-3 py-2.5 bg-indigo-50 dark:bg-indigo-900/30 border border-indigo-200 dark:border-indigo-800/50 rounded-xl text-[11px] text-indigo-700 dark:text-indigo-300">
                                                <Icon name="link" size={14} className="mt-0.5 shrink-0" />
                                                <span><strong>Sinkronisasi Otomatis Aktif:</strong> Menyimpan data ini akan otomatis membuat/memperbarui entri proyek di menu <strong>List Proyek</strong>. Status Turun/Tidak Turun tetap bisa diubah dari List Proyek.</span>
                                            </div>
                                        )}
                                        <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                                            <div>
                                                <label className="block text-xs font-bold text-slate-700 dark:text-slate-300 mb-1">Tipe Proyek</label>
                                                <select required value={formData.projectType || 'Pengawasan'} onChange={e => setFormData({ ...formData, projectType: e.target.value })} className="w-full p-3 rounded-xl border border-slate-200 dark:border-slate-700/50 focus:border-indigo-500 focus:ring-2 focus:ring-indigo-500/20 outline-none text-sm bg-slate-50/50 dark:bg-slate-900/50 text-slate-800 dark:text-slate-200 shadow-inner transition-all">
                                                    <option value="Pengawasan">Pengawasan</option>
                                                    <option value="Perencanaan">Perencanaan</option>
                                                </select>
                                            </div>
                                            <div>
                                                <label className="block text-xs font-bold text-slate-700 dark:text-slate-300 mb-1">Jenis Kontrak</label>
                                                <select required value={formData.contractType} onChange={e => setFormData({ ...formData, contractType: e.target.value })} className="w-full p-3 rounded-xl border border-slate-200 dark:border-slate-700/50 focus:border-indigo-500 focus:ring-2 focus:ring-indigo-500/20 outline-none text-sm bg-slate-50/50 dark:bg-slate-900/50 text-slate-800 dark:text-slate-200 shadow-inner transition-all">
                                                    <option value="Waktu Penugasan">Waktu Penugasan</option>
                                                    <option value="Lumsum">Lumsum</option>
                                                </select>
                                            </div>
                                            <div>
                                                <label className="block text-xs font-bold text-slate-700 dark:text-slate-300 mb-1">Jenis Tender</label>
                                                <select required value={formData.tenderType} onChange={e => setFormData({ ...formData, tenderType: e.target.value })} className="w-full p-3 rounded-xl border border-slate-200 dark:border-slate-700/50 focus:border-indigo-500 focus:ring-2 focus:ring-indigo-500/20 outline-none text-sm bg-slate-50/50 dark:bg-slate-900/50 text-slate-800 dark:text-slate-200 shadow-inner transition-all">
                                                    <option value="Tender">Tender</option>
                                                    <option value="PL">Penunjukan Langsung (PL)</option>
                                                </select>
                                            </div>
                                        </div>
                                        <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                                            <div>
                                                <label className="block text-xs font-bold text-slate-700 dark:text-slate-300 mb-1">SPMK (Tanggal)</label>
                                                <div className="relative w-full rounded-xl border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-800 focus-within:border-indigo-500 focus-within:ring-1 focus-within:ring-indigo-500 transition-colors">
                                                    <div className="absolute inset-0 p-2.5 flex items-center justify-between pointer-events-none">
                                                        <span className={`text-sm ${formData.startDate ? 'text-slate-800 dark:text-slate-200' : 'text-slate-400'}`}>
                                                            {formData.startDate ? formatDateIndo(formData.startDate) : "Pilih Tanggal"}
                                                        </span>
                                                        <Icon name="calendar" size={16} className="text-slate-400" />
                                                    </div>
                                                    <input type="date" required value={formData.startDate} onChange={e => setFormData({ ...formData, startDate: e.target.value })} className="w-full p-2.5 opacity-0 cursor-pointer outline-none" />
                                                </div>
                                            </div>
                                            <div>
                                                <label className="block text-xs font-bold text-slate-700 dark:text-slate-300 mb-1">Durasi Kontrak (Hari)</label>
                                                <input type="number" min="1" required value={formData.duration} onChange={e => setFormData({ ...formData, duration: e.target.value })} placeholder="Cth: 90" className="w-full p-3 rounded-xl border border-slate-200 dark:border-slate-700/50 focus:border-indigo-500 focus:ring-2 focus:ring-indigo-500/20 outline-none text-sm bg-slate-50/50 dark:bg-slate-900/50 text-slate-800 dark:text-slate-200 shadow-inner transition-all" />
                                            </div>
                                            <div>
                                                <label className="block text-xs font-bold text-slate-700 dark:text-slate-300 mb-1">Akhir Kontrak Induk</label>
                                                <div className="relative w-full rounded-xl border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-800 focus-within:border-indigo-500 focus-within:ring-1 focus-within:ring-indigo-500 transition-colors">
                                                    <div className="absolute inset-0 p-2.5 flex items-center justify-between pointer-events-none">
                                                        <span className={`text-sm ${endDate ? 'text-slate-800 dark:text-slate-200' : 'text-slate-400'}`}>
                                                            {endDate ? formatDateIndo(endDate) : "Pilih Tanggal"}
                                                        </span>
                                                        <Icon name="calendar" size={16} className="text-slate-400" />
                                                    </div>
                                                    <input type="date" required value={endDate} onChange={e => setEndDate(e.target.value)} className="w-full p-2.5 opacity-0 cursor-pointer outline-none" />
                                                </div>
                                            </div>
                                        </div>
                                    </div>

                                    {/* SEKSI PLOTTING TENAGA AHLI */}
                                    <div className="space-y-4 pt-4 border-t border-slate-200 dark:border-slate-700">
                                        <div className="flex justify-between items-center">
                                            <h4 className="font-bold text-slate-700 dark:text-slate-200">Plotting Tenaga Ahli</h4>
                                            <button type="button" onClick={handleAddExpertRow} className="text-xs font-bold text-indigo-600 hover:text-indigo-800 dark:text-indigo-400 bg-indigo-50 hover:bg-indigo-100 dark:bg-indigo-900/30 px-3 py-1.5 rounded-lg transition-colors flex items-center gap-1">
                                                <Icon name="plus" size={14} /> Tambah Personil
                                            </button>
                                        </div>

                                        {(formData.experts || []).length === 0 ? (
                                            <div className="p-6 text-center border border-dashed border-slate-300 dark:border-slate-700 rounded-xl bg-slate-50/50 dark:bg-slate-800/50">
                                                <p className="text-sm text-slate-500">Belum ada tenaga ahli yang diplot pada pekerjaan ini.</p>
                                            </div>
                                        ) : (
                                            <div className="space-y-3">
                                                {formData.experts.map((expPlot, idx) => {
                                                    const selectedExpertObj = experts.find(e => e.id === expPlot.expertId);
                                                    const availableCerts = selectedExpertObj?.certificates || [];

                                                    return (
                                                        <div key={idx} className="p-4 bg-slate-50 dark:bg-slate-800/50 rounded-xl border border-slate-200 dark:border-slate-700 flex flex-col md:flex-row gap-3 relative group">
                                                            <div className="flex-1">
                                                                <label className="block text-[10px] uppercase font-bold text-slate-500 mb-1">Nama</label>
                                                                <input
                                                                    type="text"
                                                                    list={`expert-list-${idx}`}
                                                                    required
                                                                    value={expPlot.expertSearchTemp !== undefined ? expPlot.expertSearchTemp : (selectedExpertObj ? selectedExpertObj.name : '')}
                                                                    onChange={e => {
                                                                        const val = e.target.value;
                                                                        const matched = experts.find(ex => ex.name === val);
                                                                        setFormData(prev => {
                                                                            const newExperts = [...prev.experts];
                                                                            if (matched) {
                                                                                newExperts[idx] = { ...newExperts[idx], expertId: matched.id, expertSearchTemp: undefined, certificateName: '', additionalCertificates: [] };
                                                                            } else {
                                                                                newExperts[idx] = { ...newExperts[idx], expertId: '', expertSearchTemp: val };
                                                                            }
                                                                            return { ...prev, experts: newExperts };
                                                                        });
                                                                    }}
                                                                    placeholder="Cari atau ketik nama..."
                                                                    className="w-full p-2.5 rounded-xl border border-slate-200 dark:border-slate-700/50 focus:ring-2 focus:ring-indigo-500/20 outline-none text-sm bg-slate-50/50 dark:bg-slate-900/50 text-slate-800 dark:text-slate-200 shadow-inner transition-all"
                                                                />
                                                                <datalist id={`expert-list-${idx}`}>
                                                                    {experts.map(e => <option key={e.id} value={e.name} />)}
                                                                </datalist>
                                                            </div>
                                                            <div className="flex-1">
                                                                <label className="block text-[10px] uppercase font-bold text-slate-500 mb-1">Peran</label>
                                                                <select required value={expPlot.role || ''} onChange={e => handleExpertRowChange(idx, 'role', e.target.value)} className="w-full p-2.5 rounded-xl border border-slate-200 dark:border-slate-700/50 focus:ring-2 focus:ring-indigo-500/20 outline-none text-sm bg-slate-50/50 dark:bg-slate-900/50 text-slate-800 dark:text-slate-200 shadow-inner transition-all">
                                                                    <option value="">-- Pilih Peran --</option>
                                                                    {((formData.projectType || 'Pengawasan') === 'Perencanaan'
                                                                        ? ['Team Leader', 'Tenaga Ahli', 'K3', 'Quantity Surveyor / Cost Estimator', 'Drafter / Operator CAD', 'Surveyor', 'Administrasi']
                                                                        : ['Team Leader', 'Tenaga Ahli', 'Inspector', 'Laboratory Technician', 'Quantity Surveyor', 'K3', 'Administrasi']
                                                                    ).map(r => (
                                                                        <option key={r} value={r}>{r}</option>
                                                                    ))}
                                                                </select>
                                                            </div>
                                                            <div className="flex-1 flex flex-col gap-2">
                                                                <div>
                                                                    <label className="block text-[10px] uppercase font-bold text-slate-500 mb-1">Sertifikat Utama</label>
                                                                    <select value={expPlot.certificateName || ''} onChange={e => handleExpertRowChange(idx, 'certificateName', e.target.value)} className="w-full p-2.5 rounded-xl border border-slate-200 dark:border-slate-700/50 focus:ring-2 focus:ring-indigo-500/20 outline-none text-sm bg-slate-50/50 dark:bg-slate-900/50 text-slate-800 dark:text-slate-200 shadow-inner transition-all">
                                                                        <option value="">-- Tidak Pakai / Kosong --</option>
                                                                        {availableCerts.map((c, i) => {
                                                                            const isExpired = c.expiredDate && new Date(c.expiredDate) < new Date();
                                                                            return (
                                                                                <option key={`cert1-${i}`} value={c.certName} disabled={isExpired} className={isExpired ? "text-slate-400" : ""}>
                                                                                    {c.certName} ({c.certLevel}) {isExpired ? '(EXPIRED)' : ''}
                                                                                </option>
                                                                            );
                                                                        })}
                                                                    </select>
                                                                </div>
                                                                {(expPlot.additionalCertificates || []).map((addCert, cIdx) => (
                                                                    <div key={`add-cert-${cIdx}`} className="flex gap-2">
                                                                        <div className="flex-1">
                                                                            <label className="block text-[10px] uppercase font-bold text-slate-500 mb-1">Sertifikat Tambahan {cIdx + 1}</label>
                                                                            <select value={addCert || ''} onChange={e => {
                                                                                const newAdditional = [...(expPlot.additionalCertificates || [])];
                                                                                newAdditional[cIdx] = e.target.value;
                                                                                handleExpertRowChange(idx, 'additionalCertificates', newAdditional);
                                                                            }} className="w-full p-2.5 rounded-xl border border-slate-200 dark:border-slate-700/50 focus:ring-2 focus:ring-indigo-500/20 outline-none text-sm bg-slate-50/50 dark:bg-slate-900/50 text-slate-800 dark:text-slate-200 shadow-inner transition-all">
                                                                                <option value="">-- Pilih Sertifikat --</option>
                                                                                {availableCerts.map((c, i) => {
                                                                                    const isExpired = c.expiredDate && new Date(c.expiredDate) < new Date();
                                                                                    return (
                                                                                        <option key={`addcert-${cIdx}-${i}`} value={c.certName} disabled={isExpired} className={isExpired ? "text-slate-400" : ""}>
                                                                                            {c.certName} ({c.certLevel}) {isExpired ? '(EXPIRED)' : ''}
                                                                                        </option>
                                                                                    );
                                                                                })}
                                                                            </select>
                                                                        </div>
                                                                        <div className="pt-5">
                                                                            <button type="button" onClick={() => {
                                                                                const newAdditional = [...(expPlot.additionalCertificates || [])];
                                                                                newAdditional.splice(cIdx, 1);
                                                                                handleExpertRowChange(idx, 'additionalCertificates', newAdditional);
                                                                            }} className="p-2 bg-red-100 text-red-600 hover:bg-red-200 rounded-lg" title="Hapus Sertifikat">
                                                                                <Icon name="trash" size={16} />
                                                                            </button>
                                                                        </div>
                                                                    </div>
                                                                ))}
                                                                <button type="button" onClick={() => {
                                                                    const newAdditional = [...(expPlot.additionalCertificates || []), ''];
                                                                    handleExpertRowChange(idx, 'additionalCertificates', newAdditional);
                                                                }} className="text-xs text-indigo-600 dark:text-indigo-400 font-bold self-start flex items-center gap-1 hover:underline">
                                                                    <Icon name="plus" size={12} /> Tambah Sertifikat Tambahan
                                                                </button>
                                                            </div>
                                                            <div className="w-full md:w-24">
                                                                <label className="block text-[10px] uppercase font-bold text-slate-500 mb-1">Man Month</label>
                                                                <input type="text" required value={expPlot.manMonth} onChange={e => handleExpertRowChange(idx, 'manMonth', e.target.value)} placeholder="Cth: 1.5" className="w-full p-2.5 rounded-xl border border-slate-200 dark:border-slate-700/50 focus:ring-2 focus:ring-indigo-500/20 outline-none text-sm bg-slate-50/50 dark:bg-slate-900/50 text-slate-800 dark:text-slate-200 shadow-inner transition-all" />
                                                            </div>

                                                            <button type="button" onClick={() => handleRemoveExpertRow(idx)} className="absolute -top-2 -right-2 p-1.5 bg-red-100 text-red-600 hover:bg-red-200 dark:bg-red-900/50 dark:text-red-400 rounded-full shadow-sm opacity-0 group-hover:opacity-100 transition-opacity" title="Hapus Baris">
                                                                <Icon name="x" size={14} />
                                                            </button>
                                                        </div>
                                                    )
                                                })}
                                            </div>
                                        )}
                                    </div>
                                </form>
                            </div>
                            <div className="p-5 border-t border-slate-100 dark:border-slate-700/50 flex justify-end gap-3 bg-slate-50 dark:bg-slate-900/50">
                                <button onClick={() => setModalConfig({ isOpen: false, type: null })} type="button" className="px-4 py-2 text-sm font-medium text-slate-600 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-700 rounded-xl">Batal</button>
                                <button form="assignmentForm" type="submit" disabled={loading} className="px-6 py-2 bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-medium rounded-xl disabled:opacity-70 flex items-center gap-2">
                                    {loading ? <Icon name="refresh-ccw" className="animate-spin" size={16} /> : <Icon name="save" size={16} />} Simpan
                                </button>
                            </div>
                        </div>
                    </div>
                );
            };

            // --- MODAL FORMULAR ---
            const ModalForm = () => {
                if (!modalConfig.isOpen) return null;
                if (modalConfig.type === 'expert' || modalConfig.type === 'expert_cert' || modalConfig.type === 'expert_tender' || modalConfig.type === 'assignment' || modalConfig.type === 'import_expert') return null;

                const isProject = modalConfig.type === 'project';
                const isInventory = modalConfig.type === 'inventory' || modalConfig.type === 'inventory-borrow' || modalConfig.type === 'inventory-return' || modalConfig.type === 'inventory-extend' || modalConfig.type === 'inventory-cart';

                const [formData, setFormData] = useState(() => {
                    const offset = new Date().getTimezoneOffset() * 60000;
                    const defaultBorrowDate = new Date(Date.now() - offset).toISOString().split('T')[0];
                    if (modalConfig.data) {
                        const data = { ...modalConfig.data, adjustmentInput: 0 };
                        if (modalConfig.mode === 'borrow' && !data.borrowDate) {
                            data.borrowDate = defaultBorrowDate;
                        }
                        if (modalConfig.type === 'project' && data.type === 'Perencanaan' && data.surveyorTeam && data.surveyorTeam.length > 0) {
                            data.team = (data.team || []).filter(m => !data.surveyorTeam.includes(m));
                        }
                        return data;
                    }
                    if (isProject) return { name: '', client: '', type: 'Perencanaan', status: 'On Progress', progress: 0, deadline: '', spmk: '', description: '', descriptionUpdatedAt: '', team: [], surveyorTeam: [], categoryDetails: {}, teamLeader: '', pengawasanDetails: {} };
                    if (isInventory) return { id: '', name: '', type: 'Alat Ukur', condition: 'Baik', status: 'Tersedia', borrower: '', borrowDate: modalConfig.mode === 'borrow-cart' ? defaultBorrowDate : '', returnDate: '', projectAssigned: '' };
                    return { name: '', role: 'Arsitek', level: 'Staff', rating: 3, manualPoints: 0, adjustmentInput: 0 };
                });

                const [searchTeam, setSearchTeam] = useState("");
                const [freelanceInput, setFreelanceInput] = useState("");

                const handleAddFreelance = (e) => {
                    e.preventDefault();
                    const name = freelanceInput.trim();
                    if (!name) return;
                    if (!(formData.team || []).includes(name)) {
                        setFormData(prev => ({
                            ...prev,
                            team: [...(prev.team || []), name]
                        }));
                    }
                    setFreelanceInput("");
                };

                const handleEditFreelance = (oldName) => {
                    const newName = window.prompt("Edit nama personil freelance:", oldName);
                    if (newName && newName.trim() !== "" && newName !== oldName) {
                        const finalName = newName.trim();
                        if ((formData.team || []).includes(finalName)) {
                            alert("Nama personil sudah ada di dalam tim.");
                            return;
                        }
                        setFormData(prev => {
                            const newTeam = (prev.team || []).map(t => t === oldName ? finalName : t);
                            const newPengawasanDetails = { ...(prev.pengawasanDetails || {}) };
                            if (newPengawasanDetails[oldName]) {
                                newPengawasanDetails[finalName] = newPengawasanDetails[oldName];
                                delete newPengawasanDetails[oldName];
                            }
                            return {
                                ...prev,
                                team: newTeam,
                                pengawasanDetails: newPengawasanDetails
                            };
                        });
                    }
                };

                const handleFreelanceKeyDown = (e) => {
                    if (e.key === 'Enter') {
                        e.preventDefault();
                        handleAddFreelance(e);
                    }
                };

                // --- ALGORITMA AI AUTO-PLOTTING ---
                const handleAutoPlotting = () => {
                    let recommendedTeam = [];
                    let recommendedLeader = "";
                    const checkPengawasan = formData.type === 'Pengawasan' || formData.type === 'Manajemen Konstruksi';

                    // Helper mencari orang terbaik untuk role tertentu
                    const findBestPerson = (roleKeyword, excludeList, requiredLevel = null) => {
                        let candidates = resources.filter(r => r.role.toLowerCase().includes(roleKeyword.toLowerCase()) && !excludeList.includes(r.name));
                        if (requiredLevel) {
                            candidates = candidates.filter(r => r.level === requiredLevel);
                        }
                        if (candidates.length === 0) return null;

                        let bestCandidate = null;
                        let highestScore = -99999;

                        candidates.forEach(cand => {
                            // Hitung Workload
                            const activeProjCount = computedProjects.filter(p => p.computedStatus !== 'Done' && p.computedStatus !== 'Pending' && p.team && p.team.includes(cand.name)).length;
                            const workload = activeProjCount * 25; // Asumsi 1 project = 25% load

                            // Hitung skor KPI
                            const isLeaderMode = cand.level === 'Team Leader';
                            const kpiResult = isLeaderMode
                                ? calculateLeaderKPI(cand, computedProjects)
                                : calculateEmployeeKPI(cand, computedProjects);

                            let score = kpiResult.score;

                            // Penalty AI untuk Workload
                            if (workload >= 100) {
                                score -= 1000; // Diskualifikasi jika full load
                            } else {
                                score -= (workload * 0.5); // Penalti proporsional beban
                            }

                            // Tambah poin dari manual rating (1-5 star)
                            score += ((cand.rating || 0) * 10);

                            if (score > highestScore) {
                                highestScore = score;
                                bestCandidate = cand.name;
                            }
                        });

                        return bestCandidate;
                    };

                    if (!checkPengawasan) {
                        // Perencanaan
                        const leader = findBestPerson('Team Leader', recommendedTeam, 'Team Leader');
                        if (leader) recommendedLeader = leader;

                        const rolesToFill = ['Arsitek', 'Struktur', 'MEP', 'Estimator', 'Drafter'];
                        rolesToFill.forEach(role => {
                            const best = findBestPerson(role, recommendedTeam);
                            if (best) recommendedTeam.push(best);
                        });
                    } else {
                        // Pengawasan (Ambil top Inspector/Site Engineer)
                        const rolesToFill = ['Inspector', 'Site Engineer', 'Quantity Surveyor'];
                        rolesToFill.forEach(role => {
                            const best = findBestPerson(role, recommendedTeam);
                            if (best) recommendedTeam.push(best);
                        });
                    }

                    setFormData(prev => ({
                        ...prev,
                        teamLeader: recommendedLeader,
                        team: recommendedTeam
                    }));
                };

                const handleChange = (e) => {
                    const { name, value } = e.target;
                    setFormData(prev => {
                        let newData = { ...prev, [name]: value };
                        if (name === 'spmk') {
                            if (newData.pengawasanDetails) {
                                let updatedDetails = { ...newData.pengawasanDetails };
                                Object.keys(updatedDetails).forEach(member => {
                                    const manMonth = updatedDetails[member].manMonth;
                                    if (value && manMonth) {
                                        const manMonthVal = parseFloat(manMonth);
                                        if (!isNaN(manMonthVal)) {
                                            const date = new Date(value);
                                            date.setDate(date.getDate() + Math.round(manMonthVal * 30) - 1);
                                            if (date.getDay() === 6) date.setDate(date.getDate() - 1);
                                            else if (date.getDay() === 0) date.setDate(date.getDate() - 2);
                                            const yyyy = date.getFullYear();
                                            const mm = String(date.getMonth() + 1).padStart(2, '0');
                                            const dd = String(date.getDate()).padStart(2, '0');
                                            updatedDetails[member] = { ...updatedDetails[member], deadline: `${yyyy}-${mm}-${dd}` };
                                        }
                                    } else if (!value) {
                                        updatedDetails[member] = { ...updatedDetails[member], deadline: '' };
                                    }
                                });
                                newData.pengawasanDetails = updatedDetails;
                            }
                        }
                        return newData;
                    });
                };

                // Handler khusus untuk detail kategori
                const handleCategoryDetailChange = (category, field, value) => {
                    setFormData(prev => ({
                        ...prev,
                        categoryDetails: {
                            ...prev.categoryDetails,
                            [category]: {
                                ...(prev.categoryDetails?.[category] || {}),
                                [field]: value
                            }
                        }
                    }));
                };

                // --- SISTEM MICRO-CHECKLIST ---
                const handleAddTask = (category, taskName) => {
                    if (!taskName.trim()) return;
                    setFormData(prev => {
                        const currentCat = prev.categoryDetails?.[category] || { progress: 0, deadline: '', tasks: [], completedTasks: [] };
                        const currentTasks = currentCat.tasks || [];
                        if (currentTasks.includes(taskName.trim())) return prev;
                        const newTasks = [...currentTasks, taskName.trim()];

                        const completed = currentCat.completedTasks || [];
                        const newProgress = newTasks.length > 0 ? Math.round((completed.length / newTasks.length) * 100) : 0;

                        return {
                            ...prev,
                            categoryDetails: {
                                ...prev.categoryDetails,
                                [category]: { ...currentCat, tasks: newTasks, progress: newProgress }
                            }
                        };
                    });
                };

                const handleRemoveTask = (category, taskName) => {
                    setFormData(prev => {
                        const currentCat = prev.categoryDetails?.[category] || { progress: 0, deadline: '', tasks: [], completedTasks: [] };
                        const newTasks = (currentCat.tasks || []).filter(t => t !== taskName);
                        const newCompleted = (currentCat.completedTasks || []).filter(t => t !== taskName);

                        const newProgress = newTasks.length > 0 ? Math.round((newCompleted.length / newTasks.length) * 100) : 0;

                        return {
                            ...prev,
                            categoryDetails: {
                                ...prev.categoryDetails,
                                [category]: { ...currentCat, tasks: newTasks, completedTasks: newCompleted, progress: newProgress }
                            }
                        };
                    });
                };

                const handleToggleTask = (category, taskName) => {
                    setFormData(prev => {
                        const currentCat = prev.categoryDetails?.[category] || { progress: 0, deadline: '', tasks: [], completedTasks: [] };
                        const currentCompleted = currentCat.completedTasks || [];
                        const isCompleted = currentCompleted.includes(taskName);
                        const newCompleted = isCompleted
                            ? currentCompleted.filter(t => t !== taskName)
                            : [...currentCompleted, taskName];

                        const totalTasks = currentCat.tasks || [];
                        const newProgress = totalTasks.length > 0 ? Math.round((newCompleted.length / totalTasks.length) * 100) : 0;

                        return {
                            ...prev,
                            categoryDetails: {
                                ...prev.categoryDetails,
                                [category]: { ...currentCat, completedTasks: newCompleted, progress: newProgress }
                            }
                        };
                    });
                };

                const handlePengawasanDetailChange = (memberName, field, value) => {
                    setFormData(prev => {
                        const currentDetail = prev.pengawasanDetails?.[memberName] || { role: 'Inspector', deadline: '', manMonth: '', statusTurun: 'Tidak Turun' };
                        let newDetail = { ...currentDetail, [field]: value };

                        if (field === 'manMonth') {
                            if (prev.spmk && newDetail.manMonth) {
                                const manMonthVal = parseFloat(newDetail.manMonth);
                                if (!isNaN(manMonthVal)) {
                                    const date = new Date(prev.spmk);
                                    date.setDate(date.getDate() + Math.round(manMonthVal * 30) - 1);
                                    if (date.getDay() === 6) date.setDate(date.getDate() - 1);
                                    else if (date.getDay() === 0) date.setDate(date.getDate() - 2);
                                    const yyyy = date.getFullYear();
                                    const mm = String(date.getMonth() + 1).padStart(2, '0');
                                    const dd = String(date.getDate()).padStart(2, '0');
                                    newDetail.deadline = `${yyyy}-${mm}-${dd}`;
                                }
                            } else if (!newDetail.manMonth) {
                                newDetail.deadline = '';
                            }
                        }

                        return {
                            ...prev,
                            pengawasanDetails: {
                                ...prev.pengawasanDetails,
                                [memberName]: newDetail
                            }
                        };
                    });
                };

                const surveyorTeamList = formData.surveyorTeam || [];
                const isPerencanaanForm = formData.type?.toLowerCase().includes('perencana');

                // Dapatkan Kategori Aktif untuk menampilkan form progres spesifik tim & menghitung rata-rata
                const activeCategories = Array.from(new Set((formData.team || []).map(memberName => {
                    if (isPerencanaanForm && surveyorTeamList.includes(memberName)) return null;
                    const res = resources.find(r => r.name === memberName);
                    return res ? getCategoryFromRole(res.role) : 'Lainnya';
                }).filter(cat => cat !== null)));

                if (surveyorTeamList.length > 0) {
                    activeCategories.push('Surveyor');
                }

                // HITUNG PROGRESS TOTAL OTOMATIS BERDASARKAN RATA-RATA SUB-TIM
                const computedTotalProgress = useMemo(() => {
                    if (!isProject || activeCategories.length === 0) return 0;
                    let total = 0;
                    activeCategories.forEach(cat => {
                        total += parseInt(formData.categoryDetails?.[cat]?.progress) || 0;
                    });
                    return Math.round(total / activeCategories.length);
                }, [formData.categoryDetails, activeCategories, isProject]);

                const handleResetTo100 = () => {
                    const dummyEmp = { ...formData };
                    let calculated = null;
                    if (formData.level === 'Team Leader') {
                        calculated = calculateLeaderKPI(dummyEmp, computedProjects);
                    } else {
                        calculated = calculateEmployeeKPI(dummyEmp, computedProjects);
                    }
                    const gap = 100 - calculated.score;
                    setFormData(prev => ({ ...prev, adjustmentInput: gap }));
                };

                const handleResetToAuto = () => {
                    const gap = -(formData.manualPoints || 0);
                    setFormData(prev => ({ ...prev, adjustmentInput: gap }));
                };


                const handleSubmit = (e) => {
                    e.preventDefault();
                    let finalPayload = { ...formData };

                    if (isInventory) {
                        if (modalConfig.mode === 'borrow-cart') {
                            handleInventoryAction('borrow-cart', {
                                selectedIds: borrowCart,
                                data: {
                                    borrower: finalPayload.borrower,
                                    borrowDate: finalPayload.borrowDate || new Date().toISOString().split('T')[0],
                                    returnDate: finalPayload.returnDate,
                                    projectAssigned: finalPayload.projectAssigned
                                }
                            });
                            return;
                        } else if (modalConfig.mode === 'borrow') {
                            finalPayload.status = 'Menunggu Verifikasi';
                            if (!finalPayload.borrowDate) finalPayload.borrowDate = new Date().toISOString().split('T')[0];
                        } else if (modalConfig.mode === 'return') {
                            finalPayload.status = 'Menunggu Verifikasi Pengembalian';
                        } else if (modalConfig.mode === 'extend') {
                            finalPayload.status = 'Menunggu Verifikasi Perpanjangan';
                            finalPayload.newReturnDate = finalPayload.returnDate;
                            finalPayload.returnDate = modalConfig.data.returnDate;
                        }
                        handleInventoryAction(modalConfig.mode, finalPayload);
                        return;
                    }

                    if (isProject) {
                        // OTOMATISASI PROGRESS PROJECT (TIDAK LAGI MANUAL)
                        finalPayload.progress = computedTotalProgress;

                        // Otomatisasi Timestamp Deskripsi: Hanya update waktu jika deskripsi berubah
                        const oldDesc = modalConfig.data ? modalConfig.data.description : '';
                        if (finalPayload.description && finalPayload.description !== oldDesc) {
                            finalPayload.descriptionUpdatedAt = new Date().toISOString();
                        } else if (!finalPayload.description) {
                            finalPayload.descriptionUpdatedAt = '';
                        } else {
                            finalPayload.descriptionUpdatedAt = modalConfig.data ? modalConfig.data.descriptionUpdatedAt : '';
                        }

                        // Bersihkan kategori detail jika member dari kategori tersebut sudah dihapus
                        const activeCats = new Set((finalPayload.team || []).map(memberName => {
                            const res = resources.find(r => r.name === memberName);
                            return res ? getCategoryFromRole(res.role) : 'Lainnya';
                        }));
                        if (finalPayload.type === 'Perencanaan' && finalPayload.surveyorTeam && finalPayload.surveyorTeam.length > 0) {
                            activeCats.add('Surveyor');
                            finalPayload.team = Array.from(new Set([...(finalPayload.team || []), ...finalPayload.surveyorTeam]));
                        }

                        const cleanDetails = {};
                        activeCats.forEach(cat => {
                            if (finalPayload.categoryDetails && finalPayload.categoryDetails[cat]) {
                                cleanDetails[cat] = finalPayload.categoryDetails[cat];
                            }
                        });

                        const teamDataObj = {
                            members: finalPayload.team || [],
                            details: cleanDetails,
                            leader: finalPayload.teamLeader || "",
                            individualStatus: finalPayload.individualStatus || {},
                            pengawasanDetails: finalPayload.pengawasanDetails || {}
                        };

                        // Enkripsi ke string agar aman masuk backend lama
                        finalPayload.team = JSON.stringify(teamDataObj).replace(/;/g, ',');
                    } else {
                        // Gabungkan role, level, dan manualPoints untuk personil
                        const newManualPoints = (finalPayload.manualPoints || 0) + (finalPayload.adjustmentInput || 0);
                        finalPayload.role = `${finalPayload.role}|${finalPayload.level}|${newManualPoints}`;
                    }
                    handleCrudAction(modalConfig.mode, modalConfig.type, finalPayload);
                };

                const filteredResources = resources.filter(res =>
                    res.name.toLowerCase().includes(searchTeam.toLowerCase()) ||
                    res.role.toLowerCase().includes(searchTeam.toLowerCase())
                );

                const isPengawasanForm = formData.type === 'Pengawasan' || formData.type === 'Manajemen Konstruksi';
                const isSyncedProject = modalConfig.mode === 'edit' && Boolean(formData.sourceAssignmentId);

                return (
                    <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[70] flex items-end sm:items-center justify-center sm:p-4 fade-in">
                        <div className={`bg-white dark:bg-slate-800 rounded-t-3xl sm:rounded-3xl shadow-xl border border-transparent dark:border-slate-700/50 w-full ${isProject ? 'sm:max-w-4xl' : 'sm:max-w-2xl'} overflow-hidden flex flex-col max-h-[90vh] sm:max-h-[95vh] pb-4 sm:pb-0`}>
                            <div className="p-5 border-b border-slate-100 dark:border-slate-700/50 flex justify-between items-center bg-slate-50 dark:bg-slate-900/50 shrink-0">
                                <h3 className="font-bold text-lg text-slate-800 dark:text-slate-100">
                                    {modalConfig.mode === 'add' ? 'Tambah ' : modalConfig.mode === 'borrow' ? 'Pinjam ' : 'Edit '}
                                    {isProject ? 'Data Proyek Terpadu' : isInventory ? 'Data Inventaris' : 'Data Personil'}
                                </h3>
                                <button onClick={closeModal} className="text-slate-400 hover:text-red-500 transition-colors"><Icon name="x" size={20} /></button>
                            </div>

                            <div className="p-6 overflow-y-auto flex-1 min-h-0">
                                <form id="crudForm" onSubmit={handleSubmit} className="space-y-5">
                                    {isProject && isSyncedProject && (
                                        <div className="flex items-start gap-2.5 px-3 py-2.5 bg-indigo-50 dark:bg-indigo-900/30 border border-indigo-200 dark:border-indigo-800/50 rounded-xl text-[11px] text-indigo-700 dark:text-indigo-300">
                                            <Icon name="link" size={14} className="mt-0.5 shrink-0" />
                                            <span><strong>Proyek Tersinkronisasi:</strong> Data utama (Nama, Klien, Tipe, Tim) dikelola otomatis melalui menu <strong>Penugasan Tenaga Ahli</strong>. Anda hanya dapat mengubah status <strong>Turun/Tidak Turun</strong> personil di sini.</span>
                                        </div>
                                    )}
                                    {isProject ? (
                                        <>
                                            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                                <div className="col-span-2 md:col-span-1">
                                                    <label className="block text-xs font-semibold text-slate-600 dark:text-slate-400 mb-1">Nama Proyek</label>
                                                    <input required name="name" value={formData.name} onChange={handleChange} disabled={isSyncedProject} className="w-full p-2.5 border border-slate-200 dark:border-slate-700 rounded-xl text-sm outline-none focus:border-blue-500 bg-slate-50 dark:bg-slate-900/50 focus:bg-white dark:focus:bg-slate-800 transition-colors dark:text-slate-200 disabled:opacity-60 disabled:cursor-not-allowed" placeholder="Cth: Perencanaan RSUD" />
                                                </div>
                                                <div className="col-span-2 md:col-span-1">
                                                    <label className="block text-xs font-semibold text-slate-600 dark:text-slate-400 mb-1">Klien / Owner</label>
                                                    <input required name="client" value={formData.client} onChange={handleChange} disabled={isSyncedProject} className="w-full p-2.5 border border-slate-200 dark:border-slate-700 rounded-xl text-sm outline-none focus:border-blue-500 bg-slate-50 dark:bg-slate-900/50 focus:bg-white dark:focus:bg-slate-800 transition-colors dark:text-slate-200 disabled:opacity-60 disabled:cursor-not-allowed" placeholder="Cth: Dinas PUPR" />
                                                </div>
                                            </div>

                                            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                                <div>
                                                    <label className="block text-xs font-semibold text-slate-600 dark:text-slate-400 mb-1">Tipe Proyek</label>
                                                    <select name="type" value={formData.type} onChange={handleChange} disabled={isSyncedProject} className="w-full p-2.5 border border-slate-200 dark:border-slate-700 rounded-xl text-sm outline-none focus:border-blue-500 bg-slate-50 dark:bg-slate-900/50 focus:bg-white dark:focus:bg-slate-800 dark:text-slate-200 disabled:opacity-60 disabled:cursor-not-allowed">
                                                        <option value="Perencanaan">Perencanaan</option>
                                                        {(modalConfig.mode === 'edit' && formData.type === 'Pengawasan') && <option value="Pengawasan">Pengawasan</option>}
                                                        <option value="Manajemen Konstruksi">Manajemen Konstruksi</option>
                                                    </select>
                                                </div>
                                                {isPerencanaanForm && (
                                                    <div>
                                                        <label className="block text-xs font-semibold text-slate-600 dark:text-slate-400 mb-1 flex items-center gap-1 text-blue-600 dark:text-blue-400">Divisi Kontrol</label>
                                                        <select name="divisiKontrol" value={formData.divisiKontrol || ""} onChange={handleChange} className="w-full p-2.5 border border-slate-200 dark:border-slate-700 rounded-xl text-sm outline-none focus:border-blue-500 bg-blue-50/50 dark:bg-blue-900/20 focus:bg-white dark:focus:bg-slate-800 dark:text-slate-200">
                                                            <option value="">-- Pilih Divisi --</option>
                                                            <option value="Divisi Jalan">Divisi Jalan</option>
                                                            <option value="Divisi Gedung">Divisi Gedung</option>
                                                            <option value="Divisi SDA">Divisi SDA</option>
                                                            <option value="Divisi Perijinan">Divisi Perijinan</option>
                                                            <option value="Divisi Tata Ruang">Divisi Tata Ruang</option>
                                                        </select>
                                                    </div>
                                                )}
                                                <div>
                                                    <label className="block text-xs font-semibold text-slate-600 dark:text-slate-400 mb-1">Status Proyek (Makro)</label>
                                                    <select name="status" value={formData.status} onChange={handleChange} className="w-full p-2.5 border border-slate-200 dark:border-slate-700 rounded-xl text-sm outline-none focus:border-blue-500 bg-slate-50 dark:bg-slate-900/50 focus:bg-white dark:focus:bg-slate-800 dark:text-slate-200">
                                                        <option value="On Progress">On Progress</option>
                                                        <option value="Done">Done</option>
                                                    </select>
                                                </div>
                                            </div>

                                            <div>
                                                <label className="block text-xs font-semibold text-slate-600 dark:text-slate-400 mb-1 flex items-center gap-1"><Icon name="file-text" size={12} /> Deskripsi / Update Proyek</label>
                                                <textarea name="description" value={formData.description || ''} onChange={handleChange} rows="3" className="w-full p-2.5 border border-slate-200 dark:border-slate-700 rounded-xl text-sm outline-none focus:border-blue-500 bg-slate-50 dark:bg-slate-900/50 focus:bg-white dark:focus:bg-slate-800 transition-colors resize-none dark:text-slate-200" placeholder="Tuliskan update terkini proyek ini... Cth: Sudah masuk tahap DED, menunggu approval dari klien untuk revisi arsitektur." />
                                            </div>

                                            <div className="flex flex-col gap-4 bg-slate-100 dark:bg-slate-800/50 p-3 rounded-xl border border-slate-200 dark:border-slate-700">
                                                {(isPengawasanForm || isPerencanaanForm) && (
                                                    <div>
                                                        <label className="block text-[11px] font-bold text-slate-700 dark:text-slate-300 mb-1">Tgl SPMK Proyek</label>
                                                        <div className="relative w-full">
                                                            <div className="flex items-center justify-between w-full p-2.5 border border-slate-200 dark:border-slate-600 rounded-xl text-sm bg-white dark:bg-slate-800 min-h-[42px] text-slate-700 dark:text-slate-200">
                                                                <span>{formData.spmk ? formatDateIndo(formData.spmk) : 'Pilih tanggal...'}</span>
                                                                <Icon name="calendar-clock" size={16} className="text-slate-400" />
                                                            </div>
                                                            <input type="date" name="spmk" className={`absolute inset-0 w-full h-full opacity-0 ${isSyncedProject ? 'cursor-not-allowed' : 'cursor-pointer'}`}
                                                                value={formData.spmk || ''}
                                                                onChange={handleChange}
                                                                disabled={isSyncedProject}
                                                                onClick={(e) => !isSyncedProject && e.target.showPicker && e.target.showPicker()}
                                                            />
                                                        </div>
                                                    </div>
                                                )}
                                                <div>
                                                    <label className="block text-[11px] font-bold text-slate-700 dark:text-slate-300 mb-1">Tenggat Waktu Kontrak Akhir</label>
                                                    <div className="relative w-full">
                                                        <div className="flex items-center justify-between w-full p-2.5 border border-slate-200 dark:border-slate-600 rounded-xl text-sm bg-white dark:bg-slate-800 min-h-[42px] text-slate-700 dark:text-slate-200">
                                                            <span>{formData.deadline ? formatDateIndo(formData.deadline) : 'Pilih tanggal...'}</span>
                                                            <Icon name="calendar-clock" size={16} className="text-slate-400" />
                                                        </div>
                                                        <input type="date" name="deadline" className={`absolute inset-0 w-full h-full opacity-0 ${isSyncedProject ? 'cursor-not-allowed' : 'cursor-pointer'}`}
                                                            value={formData.deadline || ''}
                                                            onChange={handleChange}
                                                            disabled={isSyncedProject}
                                                            onClick={(e) => !isSyncedProject && e.target.showPicker && e.target.showPicker()}
                                                        />
                                                    </div>
                                                </div>
                                                {!isPengawasanForm && (
                                                    <div>
                                                        <label className="block text-[11px] font-bold text-slate-700 dark:text-slate-300 mb-1">Progress Keseluruhan Proyek (%)</label>
                                                        <input
                                                            type="number"
                                                            readOnly
                                                            value={computedTotalProgress}
                                                            className="block w-full p-2.5 border border-slate-200 dark:border-slate-700 bg-slate-200 dark:bg-slate-700 text-slate-600 dark:text-slate-300 font-bold rounded-xl text-sm outline-none cursor-not-allowed appearance-none m-0"
                                                            title="Nilai ini dihitung otomatis"
                                                        />
                                                        <p className="text-[9px] text-slate-500 mt-1">Dihitung otomatis dari rata-rata progress sub-tim.</p>
                                                    </div>
                                                )}
                                            </div>

                                            {/* BAGIAN PENUGASAN TIM */}
                                            <div className="pt-2 border-t border-slate-200 dark:border-slate-700/50">
                                                {!isSyncedProject && (
                                                    <>
                                                        <div className="flex flex-col sm:flex-row sm:items-center justify-between mb-2 gap-2">
                                                            <label className="block text-sm font-bold text-slate-800 dark:text-slate-100">Penugasan Personil</label>
                                                            {!isPengawasanForm && (
                                                                <button type="button" onClick={handleAutoPlotting} className="flex items-center gap-1.5 px-3 py-1.5 bg-gradient-to-r from-purple-600 to-indigo-600 hover:from-purple-500 hover:to-indigo-500 text-white rounded-lg text-xs font-bold transition-all shadow-md">
                                                                    <Icon name="zap" size={14} className="text-amber-300 fill-amber-300" /> Auto-Assign AI
                                                                </button>
                                                            )}
                                                        </div>

                                                        {!isPengawasanForm && (
                                                            <div className="mb-4 bg-amber-50 dark:bg-amber-900/20 p-3 rounded-xl border border-amber-200 dark:border-amber-800/50 shadow-sm">
                                                                <label className="block text-xs font-bold text-amber-800 mb-1 flex items-center gap-1.5"><Icon name="star" size={14} className="text-amber-500" /> Pilih Team Leader Proyek</label>
                                                                <select name="teamLeader" value={formData.teamLeader || ""} onChange={handleChange} className="w-full p-2.5 border border-amber-200 dark:border-amber-700/50 rounded-xl text-sm outline-none focus:border-amber-500 bg-white dark:bg-slate-800 shadow-sm font-semibold text-slate-700 dark:text-slate-200">
                                                                    <option value="">-- Tidak Ada Team Leader --</option>
                                                                    {resources.filter(r => r.level === 'Team Leader').map(r => (
                                                                        <option key={r.id} value={r.name}>{r.name} ({r.role})</option>
                                                                    ))}
                                                                </select>
                                                            </div>
                                                        )}

                                                        <label className="block text-xs font-bold text-slate-700 dark:text-slate-300 mb-1">{isPengawasanForm ? "Daftar Personil Tim Pengawasan" : "Anggota Sub-Tim"}</label>

                                                        <div className="relative mb-3">
                                                            <div className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400">
                                                                <Icon name="search" size={14} />
                                                            </div>
                                                            <input
                                                                type="text"
                                                                placeholder="Cari nama personil untuk ditugaskan..."
                                                                className="w-full pl-8 pr-3 py-2 bg-white dark:bg-slate-800 border border-slate-300 dark:border-slate-700 rounded-lg text-sm outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100 dark:focus:ring-blue-900 transition-all shadow-sm dark:text-slate-200"
                                                                value={searchTeam}
                                                                onChange={(e) => setSearchTeam(e.target.value)}
                                                            />
                                                        </div>

                                                        <div className="border border-slate-200 dark:border-slate-700/50 rounded-xl bg-slate-50 dark:bg-slate-900/50 p-3 max-h-48 overflow-y-auto">
                                                            {resources.length === 0 ? (
                                                                <p className="text-xs text-slate-400 italic">Belum ada data tim. Tambahkan di menu Alokasi Tim.</p>
                                                            ) : filteredResources.length === 0 ? (
                                                                <p className="text-xs text-slate-400 italic">Pencarian tidak ditemukan.</p>
                                                            ) : isPengawasanForm ? (
                                                                <TeamCheckboxGroup title="Pilih Anggota Tim (Bebas Lintas Jabatan)" roleFilter={r => true} isOptional={false} filteredResources={filteredResources} formData={formData} setFormData={setFormData} />
                                                            ) : (
                                                                <>
                                                                    <TeamCheckboxGroup title="Tim Arsitek" roleFilter={r => r.role.toLowerCase().includes('arsitek')} isOptional={false} filteredResources={filteredResources} formData={formData} setFormData={setFormData} />
                                                                    <TeamCheckboxGroup title="Tim Quantity Surveyor (QS)" roleFilter={r => r.role.toLowerCase() === 'qs' || r.role.toLowerCase().includes('quantity')} isOptional={false} filteredResources={filteredResources} formData={formData} setFormData={setFormData} />
                                                                    <TeamCheckboxGroup title="Tim Struktur" roleFilter={r => r.role.toLowerCase().includes('struktur')} isOptional={true} filteredResources={filteredResources} formData={formData} setFormData={setFormData} />
                                                                    <TeamCheckboxGroup title="Tim MEP" roleFilter={r => r.role.toLowerCase().includes('mep')} isOptional={true} filteredResources={filteredResources} formData={formData} setFormData={setFormData} />
                                                                    <TeamCheckboxGroup title="Tim Tata Ruang" roleFilter={r => r.role.toLowerCase().includes('tata ruang') || r.role.toLowerCase().includes('planologi')} isOptional={true} filteredResources={filteredResources} formData={formData} setFormData={setFormData} />
                                                                    <TeamCheckboxGroup title="Lainnya" roleFilter={r => !r.role.toLowerCase().includes('arsitek') && !(r.role.toLowerCase() === 'qs' || r.role.toLowerCase().includes('quantity')) && !r.role.toLowerCase().includes('struktur') && !r.role.toLowerCase().includes('mep') && !(r.role.toLowerCase().includes('tata ruang') || r.role.toLowerCase().includes('planologi'))} isOptional={true} filteredResources={filteredResources} formData={formData} setFormData={setFormData} />
                                                                </>
                                                            )}
                                                        </div>

                                                        {!isPengawasanForm && resources.length > 0 && (
                                                            <div className="mt-4 border-t border-slate-200 dark:border-slate-700/50 pt-4">
                                                                <label className="block text-xs font-bold text-slate-700 dark:text-slate-300 mb-2">Penugasan Ekstra (Opsional)</label>
                                                                <TeamVisionaryMultiSelect title="Tim Surveyor (Sub Tim Khusus Lintas Jabatan)" filteredResources={resources} formData={formData} setFormData={setFormData} teamField="surveyorTeam" />
                                                            </div>
                                                        )}

                                                        {isPengawasanForm && (
                                                            <div className="mt-4 border-t border-slate-200 dark:border-slate-700/50 pt-4">
                                                                <label className="block text-xs font-bold text-slate-700 dark:text-slate-300 mb-1">Tambah Personil Freelance Manual</label>
                                                                <p className="text-[10px] text-slate-500 mb-2">Ketik nama personil freelance lalu tekan Enter. Personil ini akan otomatis masuk ke Rincian Penugasan di bawah namun tidak dihitung dalam KPI.</p>
                                                                <div className="flex gap-2 mb-3">
                                                                    <input
                                                                        type="text"
                                                                        value={freelanceInput}
                                                                        onChange={(e) => setFreelanceInput(e.target.value)}
                                                                        onKeyDown={handleFreelanceKeyDown}
                                                                        className="flex-1 p-2 border border-slate-200 dark:border-slate-700 rounded-lg text-sm outline-none focus:border-blue-500 bg-white dark:bg-slate-800 transition-colors dark:text-slate-200 shadow-sm"
                                                                        placeholder="Ketik nama freelance..."
                                                                    />
                                                                    <button
                                                                        type="button"
                                                                        onClick={handleAddFreelance}
                                                                        className="px-4 py-2 bg-slate-100 dark:bg-slate-700 text-slate-700 dark:text-slate-200 font-semibold rounded-lg text-sm hover:bg-slate-200 dark:hover:bg-slate-600 transition-colors"
                                                                    >
                                                                        Tambah
                                                                    </button>
                                                                </div>
                                                                {formData.team && formData.team.filter(name => !resources.some(r => r.name === name)).length > 0 && (
                                                                    <div className="flex flex-wrap gap-2">
                                                                        {formData.team.filter(name => !resources.some(r => r.name === name)).map(name => (
                                                                            <div key={name} className="flex items-center gap-1.5 px-2.5 py-1 bg-amber-50 dark:bg-amber-900/30 border border-amber-200 dark:border-amber-700 rounded-md text-xs font-semibold text-amber-700 dark:text-amber-400">
                                                                                <span>{name}</span>
                                                                                <button type="button" onClick={() => handleEditFreelance(name)} className="text-amber-500 hover:text-amber-700 transition-colors ml-1" title="Edit Nama">
                                                                                    <Icon name="edit" size={12} />
                                                                                </button>
                                                                                <button type="button" onClick={() => {
                                                                                    setFormData(prev => ({
                                                                                        ...prev,
                                                                                        team: prev.team.filter(t => t !== name)
                                                                                    }))
                                                                                }} className="text-amber-500 hover:text-amber-700 transition-colors" title="Hapus">
                                                                                    <Icon name="x" size={12} />
                                                                                </button>
                                                                            </div>
                                                                        ))}
                                                                    </div>
                                                                )}
                                                            </div>
                                                        )}
                                                    </>
                                                )}

                                                {isPengawasanForm && (formData.team || []).length > 0 && (
                                                    <div className="pt-4 mt-4 border-t border-slate-100 dark:border-slate-700/50">
                                                        <label className="flex items-center gap-1.5 text-sm font-bold text-slate-800 dark:text-slate-100 mb-3"><Icon name="user-check" size={16} className="text-emerald-600 dark:text-emerald-400" /> Rincian Penugasan Personil Pengawasan</label>
                                                        <div className="space-y-2 border border-slate-200 dark:border-slate-700/50 rounded-xl p-3 bg-emerald-50/30 dark:bg-emerald-900/10">
                                                            {formData.team.map(member => {
                                                                const detail = formData.pengawasanDetails?.[member] || { role: 'Inspector', deadline: '', manMonth: '', spmk: '', statusTurun: 'Tidak Turun' };
                                                                return (
                                                                    <div key={member} className="bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 p-3 rounded-xl flex flex-col gap-3 shadow-sm">
                                                                        <div className="border-b border-slate-100 dark:border-slate-700 pb-2 flex justify-between items-center">
                                                                            <h4 className="text-sm font-bold text-slate-700 dark:text-slate-200">{member}</h4>
                                                                            <button type="button" onClick={() => {
                                                                                setFormData(prev => ({
                                                                                    ...prev,
                                                                                    team: (prev.team || []).filter(t => t !== member)
                                                                                }))
                                                                            }} className="text-xs text-red-500 hover:text-red-700 flex items-center gap-1 font-semibold transition-colors bg-red-50 dark:bg-red-900/30 px-2 py-1 rounded-md border border-red-100 dark:border-red-800/50">
                                                                                <Icon name="trash-2" size={12} /> Hapus
                                                                            </button>
                                                                        </div>
                                                                        <div className="grid grid-cols-1 sm:grid-cols-3 lg:grid-cols-[4fr_3fr_3fr] gap-3">
                                                                            <div className="flex flex-col justify-end">
                                                                                <label className="block text-[9px] font-bold text-slate-500 dark:text-slate-400 mb-1 uppercase">Peran / Jabatan</label>
                                                                                <select className="block w-full p-2 border border-slate-200 dark:border-slate-700 rounded-lg text-xs outline-none focus:border-emerald-500 bg-slate-50 dark:bg-slate-900/50 focus:bg-white dark:focus:bg-slate-800 transition-colors dark:text-slate-200 disabled:opacity-60 disabled:cursor-not-allowed"
                                                                                    value={detail.role || 'Inspector'}
                                                                                    onChange={(e) => handlePengawasanDetailChange(member, 'role', e.target.value)}
                                                                                    disabled={isSyncedProject}
                                                                                >
                                                                                    <option value="Team Leader">Team Leader</option>
                                                                                    <option value="Tenaga Ahli">Tenaga Ahli</option>
                                                                                    <option value="Inspector">Inspector</option>
                                                                                    <option value="Laboratory Technician">Laboratory Technician</option>
                                                                                    <option value="Quantity Surveyor">Quantity Surveyor</option>
                                                                                    <option value="K3">K3</option>
                                                                                    <option value="Administrasi">Administrasi</option>
                                                                                </select>
                                                                            </div>
                                                                            <div className="flex flex-col justify-end">
                                                                                <label className="block text-[9px] font-bold text-slate-500 dark:text-slate-400 mb-1 uppercase whitespace-nowrap">Status</label>
                                                                                <select className="block w-full p-2 border border-slate-200 dark:border-slate-700 rounded-lg text-xs outline-none focus:border-emerald-500 bg-slate-50 dark:bg-slate-900/50 focus:bg-white dark:focus:bg-slate-800 transition-colors dark:text-slate-200"
                                                                                    value={detail.statusTurun || 'Tidak Turun'}
                                                                                    onChange={(e) => handlePengawasanDetailChange(member, 'statusTurun', e.target.value)}
                                                                                >
                                                                                    <option value="Turun">Turun</option>
                                                                                    <option value="Tidak Turun">Tidak Turun</option>
                                                                                    <option value="Belum Diketahui">Belum Diketahui</option>
                                                                                </select>
                                                                            </div>
                                                                            <div className="flex flex-col justify-end">
                                                                                <label className="block text-[9px] font-bold text-slate-500 dark:text-slate-400 mb-1 uppercase whitespace-nowrap">Lama (Man/Month)</label>
                                                                                <input type="number" step="0.1" className="block w-full p-2 border border-slate-200 dark:border-slate-700 rounded-lg text-xs outline-none focus:border-emerald-500 bg-slate-50 dark:bg-slate-900/50 focus:bg-white dark:focus:bg-slate-800 transition-colors dark:text-slate-200 disabled:opacity-60 disabled:cursor-not-allowed"
                                                                                    value={detail.manMonth || ''}
                                                                                    onChange={(e) => handlePengawasanDetailChange(member, 'manMonth', e.target.value)}
                                                                                    placeholder="Cth: 1.5"
                                                                                    disabled={isSyncedProject}
                                                                                />
                                                                            </div>
                                                                            
                                                                        </div>
                                                                    </div>
                                                                )
                                                            })}
                                                        </div>
                                                    </div>
                                                )}
                                            </div>

                                            {/* BAGIAN TARGET MICRO PER SUB-TIM */}
                                            {!isPengawasanForm && activeCategories.length > 0 && (
                                                <div className="pt-4 border-t border-slate-100 dark:border-slate-700/50">
                                                    <label className="flex items-center gap-1.5 text-sm font-bold text-slate-800 dark:text-slate-100 mb-3"><Icon name="target" size={16} className="text-blue-600 dark:text-blue-400" /> Target Khusus Per Sub-Tim</label>
                                                    <div className="space-y-2 border border-slate-200 dark:border-slate-700/50 rounded-xl p-3 bg-blue-50/30 dark:bg-blue-900/10">
                                                        {['Arsitek', 'QS', 'Struktur', 'MEP', 'Tata Ruang', 'Surveyor', 'Lainnya'].map(cat => {
                                                            if (!activeCategories.includes(cat)) return null;

                                                            const details = formData.categoryDetails?.[cat] || { progress: 0, deadline: '' };
                                                            const catMembers = cat === 'Surveyor' ? (formData.surveyorTeam || []) : formData.team.filter(m => {
                                                                if (isPerencanaanForm && (formData.surveyorTeam || []).includes(m)) return false;
                                                                const r = resources.find(x => x.name === m);
                                                                return r && getCategoryFromRole(r.role) === cat;
                                                            });

                                                            return (
                                                                <div key={cat} className="bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 p-3 rounded-xl flex flex-col gap-3 shadow-sm">
                                                                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                                                                        <div className="w-full sm:w-1/4 sm:border-r border-slate-100 dark:border-slate-700 sm:pr-2">
                                                                            <h4 className="text-sm font-bold text-slate-700 dark:text-slate-200 flex items-center gap-1.5"><Icon name="target" size={14} className="text-blue-500" /> {cat}</h4>
                                                                            <p className="text-[9px] text-slate-500 truncate mt-0.5" title={catMembers.join(', ')}>
                                                                                {catMembers.join(', ')}
                                                                            </p>
                                                                        </div>
                                                                        <div className="w-full sm:w-1/4">
                                                                            {(details.tasks || []).length > 0 ? (
                                                                                <>
                                                                                    <label className="block text-[9px] font-bold text-slate-500 mb-1 uppercase">Progress Terkunci (%)</label>
                                                                                    <div className="flex items-center gap-2">
                                                                                        <div className="flex-1 h-2 bg-slate-200 dark:bg-slate-700 rounded-full overflow-hidden">
                                                                                            <div className="h-full bg-blue-500 transition-all duration-500" style={{ width: `${details.progress || 0}%` }}></div>
                                                                                        </div>
                                                                                        <span className="text-xs font-bold text-slate-700 dark:text-slate-200">{details.progress || 0}%</span>
                                                                                    </div>
                                                                                </>
                                                                            ) : (
                                                                                <>
                                                                                    <div className="flex items-center justify-between mb-1">
                                                                                        <label className="block text-[9px] font-bold text-slate-500 uppercase">Progress Manual</label>
                                                                                        <span className="text-[10px] font-bold text-blue-600 dark:text-blue-400">{details.progress || 0}%</span>
                                                                                    </div>
                                                                                    <div className="relative w-full h-[34px] flex items-center">
                                                                                        <div className="relative w-full h-2 bg-slate-200 dark:bg-slate-700 rounded-lg flex items-center">
                                                                                            <div className="absolute top-0 left-0 h-full bg-green-500 rounded-lg pointer-events-none" style={{ width: `${details.progress || 0}%` }}></div>
                                                                                            <div className="absolute w-4 h-4 bg-white border-[3px] border-green-500 rounded-full shadow pointer-events-none" style={{ left: `calc(${details.progress || 0}% - 8px)` }}></div>
                                                                                            <input
                                                                                                type="range"
                                                                                                min="0"
                                                                                                max="100"
                                                                                                value={details.progress || 0}
                                                                                                onChange={(e) => handleCategoryDetailChange(cat, 'progress', parseInt(e.target.value))}
                                                                                                className="absolute top-0 left-0 w-full h-full opacity-0 cursor-pointer m-0 z-10"
                                                                                            />
                                                                                        </div>
                                                                                    </div>
                                                                                </>
                                                                            )}
                                                                        </div>
                                                                        <div className="w-full sm:w-1/4 sm:border-r border-slate-100 dark:border-slate-700 sm:pr-2">
                                                                            <label className="block text-[9px] font-bold text-slate-500 mb-1 uppercase">Tanggal Mulai</label>
                                                                            <div className="relative w-full">
                                                                                <div className="flex items-center justify-between w-full p-2 border border-slate-200 dark:border-slate-700 rounded-lg text-xs bg-slate-50 dark:bg-slate-900/50 focus-within:border-blue-500 focus-within:bg-white dark:focus-within:bg-slate-800 transition-colors min-h-[34px] text-slate-700 dark:text-slate-200">
                                                                                    <span>{details.startDate ? formatDateIndo(details.startDate) : 'Pilih...'}</span>
                                                                                    <Icon name="calendar-clock" size={14} className="text-slate-400" />
                                                                                </div>
                                                                                <input type="date" className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                                                                                    value={details.startDate || ''}
                                                                                    onChange={(e) => handleCategoryDetailChange(cat, 'startDate', e.target.value)}
                                                                                    onClick={(e) => e.target.showPicker && e.target.showPicker()}
                                                                                />
                                                                            </div>
                                                                        </div>
                                                                        <div className="w-full sm:w-1/4">
                                                                            <label className="block text-[9px] font-bold text-slate-500 mb-1 uppercase">Deadline Tim</label>
                                                                            <div className="relative w-full">
                                                                                <div className="flex items-center justify-between w-full p-2 border border-slate-200 dark:border-slate-700 rounded-lg text-xs bg-slate-50 dark:bg-slate-900/50 focus-within:border-blue-500 focus-within:bg-white dark:focus-within:bg-slate-800 transition-colors min-h-[34px] text-slate-700 dark:text-slate-200">
                                                                                    <span>{details.deadline ? formatDateIndo(details.deadline) : 'Pilih tanggal...'}</span>
                                                                                    <Icon name="calendar-clock" size={14} className="text-slate-400" />
                                                                                </div>
                                                                                <input type="date" className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                                                                                    value={details.deadline || ''}
                                                                                    onChange={(e) => handleCategoryDetailChange(cat, 'deadline', e.target.value)}
                                                                                    onClick={(e) => e.target.showPicker && e.target.showPicker()}
                                                                                />
                                                                            </div>
                                                                        </div>
                                                                    </div>

                                                                    {/* Micro-Tasks Checklist */}
                                                                    <div className="bg-slate-50 dark:bg-slate-900/50 p-2.5 rounded-lg border border-slate-100 dark:border-slate-700/50 mt-1">
                                                                        <label className="block text-[10px] font-bold text-slate-500 dark:text-slate-400 mb-2 uppercase">Custom Micro-Tasks</label>
                                                                        <input
                                                                            type="text"
                                                                            placeholder={
                                                                                cat === 'QS' ? "Ketik tugas baru (Cth: RAB) lalu tekan Enter..." :
                                                                                    cat === 'Struktur' ? "Ketik tugas baru (Cth: Perhitungan Struktur Kolom) lalu tekan Enter..." :
                                                                                        cat === 'MEP' ? "Ketik tugas baru (Cth: Perhitungan sanitair) lalu tekan Enter..." :
                                                                                            "Ketik tugas baru (Cth: Denah Lantai 1) lalu tekan Enter..."
                                                                            }
                                                                            className="w-full p-2 text-xs border border-slate-200 dark:border-slate-700 rounded-md mb-2 bg-white dark:bg-slate-800 outline-none focus:border-blue-500 dark:text-slate-200 transition-colors"
                                                                            onKeyDown={(e) => {
                                                                                if (e.key === 'Enter') {
                                                                                    e.preventDefault();
                                                                                    handleAddTask(cat, e.target.value);
                                                                                    e.target.value = '';
                                                                                }
                                                                            }}
                                                                        />
                                                                        <div className="flex flex-col gap-1.5 max-h-32 overflow-y-auto pr-1">
                                                                            {(details.tasks || []).length === 0 ? (
                                                                                <p className="text-[10px] text-slate-400 italic">Belum ada tugas. Tambahkan tugas agar progres dapat dihitung.</p>
                                                                            ) : (
                                                                                (details.tasks || []).map((task, idx) => {
                                                                                    const isChecked = (details.completedTasks || []).includes(task);
                                                                                    return (
                                                                                        <div key={idx} className={`flex items-center justify-between p-1.5 rounded-md border text-xs transition-all ${isChecked ? 'bg-blue-50 border-blue-200 dark:bg-blue-900/20 dark:border-blue-800/50' : 'bg-white border-slate-200 dark:bg-slate-800 dark:border-slate-700'}`}>
                                                                                            <label className="flex items-center gap-2 cursor-pointer flex-1">
                                                                                                <input type="checkbox" className="rounded border-slate-300 text-blue-600 focus:ring-blue-500 cursor-pointer"
                                                                                                    checked={isChecked}
                                                                                                    onChange={() => handleToggleTask(cat, task)}
                                                                                                />
                                                                                                <span className={`${isChecked ? 'text-blue-800 dark:text-blue-300 line-through opacity-70' : 'text-slate-700 dark:text-slate-200 font-medium'}`}>{task}</span>
                                                                                            </label>
                                                                                            <button type="button" onClick={() => handleRemoveTask(cat, task)} className="text-slate-400 hover:text-red-500 transition-colors px-1 ml-2 flex-shrink-0" title="Hapus Tugas">
                                                                                                <Icon name="x" size={12} />
                                                                                            </button>
                                                                                        </div>
                                                                                    )
                                                                                })
                                                                            )}
                                                                        </div>
                                                                    </div>
                                                                </div>
                                                            )
                                                        })}
                                                    </div>
                                                </div>
                                            )}
                                        </>
                                    ) : isInventory ? (
                                        modalConfig.mode === 'borrow' ? (
                                            <div className="grid grid-cols-1 gap-4">
                                                <div>
                                                    <label className="block text-xs font-semibold text-slate-600 dark:text-slate-400 mb-1">Peminjam (Team Leader)</label>
                                                    <select required name="borrower" value={formData.borrower} onChange={handleChange} className="w-full p-2.5 border border-slate-200 dark:border-slate-700 rounded-xl text-sm outline-none focus:border-blue-500 bg-slate-50 dark:bg-slate-900/50 focus:bg-white dark:focus:bg-slate-800 transition-colors dark:text-slate-200">
                                                        <option value="">-- Pilih Team Leader --</option>
                                                        {resources.filter(r => r.level === 'Team Leader').map(r => (
                                                            <option key={r.name} value={r.name}>{r.name}</option>
                                                        ))}
                                                    </select>
                                                </div>
                                                <div>
                                                    <label className="block text-xs font-semibold text-slate-600 dark:text-slate-400 mb-1">Tanggal Dipinjam</label>
                                                    <input type="date" required name="borrowDate" value={formData.borrowDate} onChange={handleChange} className="w-full p-2.5 border border-slate-200 dark:border-slate-700 rounded-xl text-sm outline-none focus:border-blue-500 bg-slate-50 dark:bg-slate-900/50 focus:bg-white dark:focus:bg-slate-800 transition-colors dark:text-slate-200" />
                                                </div>
                                                <div>
                                                    <label className="block text-xs font-semibold text-slate-600 dark:text-slate-400 mb-1">Tanggal Pengembalian</label>
                                                    <input type="date" required name="returnDate" value={formData.returnDate} onChange={handleChange} className="w-full p-2.5 border border-slate-200 dark:border-slate-700 rounded-xl text-sm outline-none focus:border-blue-500 bg-slate-50 dark:bg-slate-900/50 focus:bg-white dark:focus:bg-slate-800 transition-colors dark:text-slate-200" />
                                                </div>
                                                <div>
                                                    <label className="block text-xs font-semibold text-slate-600 dark:text-slate-400 mb-1">Proyek (Opsional)</label>
                                                    <input name="projectAssigned" value={formData.projectAssigned} onChange={handleChange} className="w-full p-2.5 border border-slate-200 dark:border-slate-700 rounded-xl text-sm outline-none focus:border-blue-500 bg-slate-50 dark:bg-slate-900/50 focus:bg-white dark:focus:bg-slate-800 transition-colors dark:text-slate-200" placeholder="Cth: Perencanaan RSUD" />
                                                </div>
                                            </div>
                                        ) : modalConfig.mode === 'borrow-cart' ? (
                                            <div className="grid grid-cols-1 gap-4">
                                                <div className="p-3 bg-indigo-50 dark:bg-indigo-900/20 border border-indigo-100 dark:border-indigo-800/50 rounded-xl mb-2">
                                                    <p className="text-xs text-indigo-800 dark:text-indigo-300">
                                                        Anda akan meminjam <strong className="font-bold">{borrowCart.length} alat sekaligus</strong>. Data formulir ini akan diaplikasikan ke semua alat yang Anda pilih.
                                                    </p>
                                                </div>
                                                <div>
                                                    <label className="block text-xs font-semibold text-slate-600 dark:text-slate-400 mb-1">Peminjam (Team Leader)</label>
                                                    <select required name="borrower" value={formData.borrower} onChange={handleChange} className="w-full p-2.5 border border-slate-200 dark:border-slate-700 rounded-xl text-sm outline-none focus:border-blue-500 bg-slate-50 dark:bg-slate-900/50 focus:bg-white dark:focus:bg-slate-800 transition-colors dark:text-slate-200">
                                                        <option value="">-- Pilih Team Leader --</option>
                                                        {resources.filter(r => r.level === 'Team Leader').map(r => (
                                                            <option key={r.name} value={r.name}>{r.name}</option>
                                                        ))}
                                                    </select>
                                                </div>
                                                <div>
                                                    <label className="block text-xs font-semibold text-slate-600 dark:text-slate-400 mb-1">Tanggal Dipinjam</label>
                                                    <input type="date" required name="borrowDate" value={formData.borrowDate} onChange={handleChange} className="w-full p-2.5 border border-slate-200 dark:border-slate-700 rounded-xl text-sm outline-none focus:border-blue-500 bg-slate-50 dark:bg-slate-900/50 focus:bg-white dark:focus:bg-slate-800 transition-colors dark:text-slate-200" />
                                                </div>
                                                <div>
                                                    <label className="block text-xs font-semibold text-slate-600 dark:text-slate-400 mb-1">Tanggal Pengembalian</label>
                                                    <input type="date" required name="returnDate" value={formData.returnDate} onChange={handleChange} className="w-full p-2.5 border border-slate-200 dark:border-slate-700 rounded-xl text-sm outline-none focus:border-blue-500 bg-slate-50 dark:bg-slate-900/50 focus:bg-white dark:focus:bg-slate-800 transition-colors dark:text-slate-200" />
                                                </div>
                                                <div>
                                                    <label className="block text-xs font-semibold text-slate-600 dark:text-slate-400 mb-1">Proyek (Opsional)</label>
                                                    <input name="projectAssigned" value={formData.projectAssigned} onChange={handleChange} className="w-full p-2.5 border border-slate-200 dark:border-slate-700 rounded-xl text-sm outline-none focus:border-blue-500 bg-slate-50 dark:bg-slate-900/50 focus:bg-white dark:focus:bg-slate-800 transition-colors dark:text-slate-200" placeholder="Cth: Perencanaan RSUD" />
                                                </div>
                                            </div>
                                        ) : modalConfig.mode === 'extend' ? (
                                            <div className="grid grid-cols-1 gap-4">
                                                <div className="p-3 bg-amber-50 dark:bg-amber-900/20 border border-slate-200 dark:border-amber-800/50 rounded-xl mb-2">
                                                    <p className="text-xs text-amber-800 dark:text-amber-300">
                                                        Perpanjang masa pinjam untuk alat <strong className="font-bold">{formData.name}</strong> yang sedang dipinjam oleh <strong className="font-bold">{formData.borrower}</strong>.
                                                    </p>
                                                </div>
                                                <div>
                                                    <label className="block text-xs font-semibold text-slate-600 dark:text-slate-400 mb-1">Tanggal Pengembalian Baru</label>
                                                    <input type="date" required name="returnDate" value={formData.returnDate} onChange={handleChange} className="w-full p-2.5 border border-slate-200 dark:border-slate-700 rounded-xl text-sm outline-none focus:border-blue-500 bg-slate-50 dark:bg-slate-900/50 focus:bg-white dark:focus:bg-slate-800 transition-colors dark:text-slate-200" />
                                                </div>
                                            </div>
                                        ) : modalConfig.mode === 'return' ? (
                                            <div className="grid grid-cols-1 gap-4">
                                                <div className="p-3 bg-indigo-50 dark:bg-indigo-900/20 border border-indigo-100 dark:border-indigo-800/50 rounded-xl mb-2">
                                                    <p className="text-xs text-indigo-800 dark:text-indigo-300">
                                                        Anda akan mengembalikan alat <strong className="font-bold">{formData.name}</strong> yang dipinjam oleh <strong className="font-bold">{formData.borrower}</strong>. Silakan perbarui kondisi terakhir alat ini.
                                                    </p>
                                                </div>
                                                <div>
                                                    <label className="block text-xs font-semibold text-slate-600 dark:text-slate-400 mb-1">Kondisi Alat Saat Dikembalikan</label>
                                                    <select name="condition" value={formData.condition} onChange={handleChange} className="w-full p-2.5 border border-slate-200 dark:border-slate-700 rounded-xl text-sm outline-none focus:border-blue-500 bg-slate-50 dark:bg-slate-900/50 focus:bg-white dark:focus:bg-slate-800 transition-colors dark:text-slate-200">
                                                        <option value="Baik">Baik</option>
                                                        <option value="Rusak Ringan">Rusak Ringan</option>
                                                        <option value="Rusak Berat">Rusak Berat</option>
                                                    </select>
                                                </div>
                                            </div>
                                        ) : (
                                            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                                {modalConfig.mode !== 'edit' && (
                                                    <div className="col-span-2 sm:col-span-1">
                                                        <label className="block text-xs font-semibold text-slate-600 dark:text-slate-400 mb-1">ID / Kode Alat</label>
                                                        <input required name="id" value={formData.id} onChange={handleChange} className="w-full p-2.5 border border-slate-200 dark:border-slate-700 rounded-xl text-sm outline-none focus:border-blue-500 bg-slate-50 dark:bg-slate-900/50 focus:bg-white dark:focus:bg-slate-800 transition-colors dark:text-slate-200" placeholder="Cth: INV-2023-001" />
                                                    </div>
                                                )}
                                                <div className={`col-span-2 ${modalConfig.mode !== 'edit' ? 'sm:col-span-1' : ''}`}>
                                                    <label className="block text-xs font-semibold text-slate-600 dark:text-slate-400 mb-1">Nama Alat</label>
                                                    <input required name="name" value={formData.name} onChange={handleChange} className="w-full p-2.5 border border-slate-200 dark:border-slate-700 rounded-xl text-sm outline-none focus:border-blue-500 bg-slate-50 dark:bg-slate-900/50 focus:bg-white dark:focus:bg-slate-800 transition-colors dark:text-slate-200" placeholder="Cth: Drone DJI Mavic" />
                                                </div>
                                                <div>
                                                    <label className="block text-xs font-semibold text-slate-600 dark:text-slate-400 mb-1">Kategori</label>
                                                    <select name="type" value={formData.type} onChange={handleChange} className="w-full p-2.5 border border-slate-200 dark:border-slate-700 rounded-xl text-sm outline-none focus:border-blue-500 bg-slate-50 dark:bg-slate-900/50 focus:bg-white dark:focus:bg-slate-800 transition-colors dark:text-slate-200">
                                                        <option value="Alat Ukur">Alat Ukur</option>
                                                        <option value="Kendaraan">Kendaraan</option>
                                                        <option value="Elektronik">Elektronik</option>
                                                        <option value="Lainnya">Lainnya</option>
                                                    </select>
                                                </div>
                                                <div>
                                                    <label className="block text-xs font-semibold text-slate-600 dark:text-slate-400 mb-1">Kondisi</label>
                                                    <select name="condition" value={formData.condition} onChange={handleChange} className="w-full p-2.5 border border-slate-200 dark:border-slate-700 rounded-xl text-sm outline-none focus:border-blue-500 bg-slate-50 dark:bg-slate-900/50 focus:bg-white dark:focus:bg-slate-800 transition-colors dark:text-slate-200">
                                                        <option value="Baik">Baik</option>
                                                        <option value="Rusak Ringan">Rusak Ringan</option>
                                                        <option value="Rusak Berat">Rusak Berat</option>
                                                    </select>
                                                </div>
                                            </div>
                                        )
                                    ) : (
                                        <>
                                            <div><label className="block text-xs font-semibold text-slate-600 dark:text-slate-400 mb-1">Nama Personil</label><input required name="name" value={formData.name} onChange={handleChange} className="w-full p-2.5 border border-slate-200 dark:border-slate-700 rounded-xl text-sm outline-none focus:border-blue-500 bg-slate-50 dark:bg-slate-900/50 focus:bg-white dark:focus:bg-slate-800 transition-colors dark:text-slate-200" placeholder="Cth: Ir. Budi Santoso" /></div>
                                            <div>
                                                <label className="block text-xs font-semibold text-slate-600 dark:text-slate-400 mb-1">Tim</label>
                                                <select name="role" value={formData.role} onChange={handleChange} className="w-full p-2.5 border border-slate-200 dark:border-slate-700 rounded-xl text-sm outline-none focus:border-blue-500 bg-slate-50 dark:bg-slate-900/50 focus:bg-white dark:focus:bg-slate-800 transition-colors dark:text-slate-200">
                                                    <option value="Arsitek">Arsitek</option>
                                                    <option value="QS">Quantity Surveyor (QS)</option>
                                                    <option value="Struktur">Struktur</option>
                                                    <option value="MEP">MEP</option>
                                                    <option value="Tata Ruang">Tata Ruang</option>
                                                    <option value="Lainnya">Lainnya</option>
                                                </select>
                                            </div>
                                            <div>
                                                <label className="block text-xs font-semibold text-slate-600 dark:text-slate-400 mb-1">Tingkat Jabatan</label>
                                                <select name="level" value={formData.level} onChange={handleChange} className="w-full p-2.5 border border-slate-200 dark:border-slate-700 rounded-xl text-sm outline-none focus:border-blue-500 bg-slate-50 dark:bg-slate-900/50 focus:bg-white dark:focus:bg-slate-800 transition-colors dark:text-slate-200">
                                                    <option value="Staff">Staff (Anggota)</option>
                                                    <option value="Team Leader">Team Leader</option>
                                                    <option value="PIC">PIC</option>
                                                    <option value="Kordinator Divisi Jalan">Kordinator Divisi Jalan</option>
                                                    <option value="Kordinator Divisi Gedung">Kordinator Divisi Gedung</option>
                                                    <option value="Kordinator Divisi SDA">Kordinator Divisi SDA</option>
                                                    <option value="Kordinator Divisi Perijinan">Kordinator Divisi Perijinan</option>
                                                    <option value="Kordinator Divisi Tata Ruang">Kordinator Divisi Tata Ruang</option>
                                                </select>
                                            </div>
                                            <div>
                                                <label className="block text-xs font-semibold text-slate-600 dark:text-slate-400 mb-1">Rating Kinerja (1-5)</label>
                                                <div className="flex gap-2 items-center p-1">
                                                    {[1, 2, 3, 4, 5].map(star => (
                                                        <button
                                                            key={star}
                                                            type="button"
                                                            onClick={() => setFormData({ ...formData, rating: star })}
                                                            className={`transition-colors focus:outline-none ${star <= (formData.rating || 3) ? 'text-amber-500' : 'text-slate-200 dark:text-slate-700'}`}
                                                        >
                                                            <Icon name="star" size={28} />
                                                        </button>
                                                    ))}
                                                    <span className="text-xs text-slate-500 dark:text-slate-400 ml-2 font-medium">Bintang {formData.rating || 3}</span>
                                                </div>
                                            </div>

                                            <div className="p-4 bg-slate-50 dark:bg-slate-900/50 border border-slate-200 dark:border-slate-700 rounded-xl space-y-3">
                                                <div>
                                                    <label className="block text-xs font-bold text-slate-700 dark:text-slate-300 mb-1">Tambah / Kurang Skor KPI (+/-)</label>
                                                    <p className="text-[10px] text-slate-500 mb-2 leading-relaxed">Nilai ini akan langsung ditambahkan secara relatif ke skor terakhir. Contoh: Jika skor 60 dan Anda isi 40, skor menjadi 100.</p>
                                                    <div className="flex flex-wrap gap-3 items-center mt-2">
                                                        <input
                                                            type="number"
                                                            name="adjustmentInput"
                                                            value={formData.adjustmentInput === '' || isNaN(formData.adjustmentInput) ? '' : (formData.adjustmentInput !== undefined ? formData.adjustmentInput : 0)}
                                                            onChange={(e) => setFormData({ ...formData, adjustmentInput: e.target.value === '' ? '' : parseInt(e.target.value) })}
                                                            className="w-24 p-2 border border-slate-300 dark:border-slate-600 rounded-lg text-sm text-center outline-none focus:border-blue-500 bg-white dark:bg-slate-800 dark:text-slate-200"
                                                        />
                                                        <button
                                                            type="button"
                                                            onClick={handleResetTo100}
                                                            className="px-3 py-2 bg-blue-100 dark:bg-blue-900/30 hover:bg-blue-200 dark:hover:bg-blue-900/50 text-blue-700 dark:text-blue-400 border border-blue-200 dark:border-blue-800/50 text-xs font-bold rounded-lg transition-colors flex items-center gap-1"
                                                        >
                                                            <Icon name="target" size={12} />
                                                            Reset ke 100
                                                        </button>
                                                        <button
                                                            type="button"
                                                            onClick={handleResetToAuto}
                                                            className="px-3 py-2 bg-slate-200 dark:bg-slate-700 hover:bg-slate-300 dark:hover:bg-slate-600 text-slate-700 dark:text-slate-300 border border-slate-300 dark:border-slate-600 text-xs font-bold rounded-lg transition-colors flex items-center gap-1"
                                                            title="Hapus semua penyesuaian manual dan kembalikan ke perhitungan murni otomatis proyek."
                                                        >
                                                            <Icon name="refresh-ccw" size={12} />
                                                            Hapus Manual
                                                        </button>
                                                    </div>
                                                </div>
                                            </div>
                                            <div className="p-3 bg-blue-50 dark:bg-blue-900/20 border border-blue-100 dark:border-blue-800/50 rounded-xl mt-4">
                                                <p className="text-xs text-blue-800 dark:text-blue-300 font-medium leading-relaxed">
                                                    💡 <strong className="font-bold">Info Cerdas:</strong> Jumlah proyek dan Persentase Beban Kerja orang ini tidak perlu diinput manual. Sistem akan otomatis menghitungnya berdasarkan seberapa banyak proyek yang ia tangani di menu List Proyek.
                                                </p>
                                            </div>
                                        </>
                                    )}
                                </form>
                            </div>

                            <div className="p-5 border-t border-slate-100 dark:border-slate-700/50 flex justify-end gap-3 bg-slate-50 dark:bg-slate-900/50 shrink-0">
                                <button onClick={closeModal} type="button" className="px-4 py-2 text-sm font-medium text-slate-600 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-700 rounded-xl transition-colors">Batal</button>
                                <button form="crudForm" type="submit" disabled={loading} className="flex items-center gap-2 px-6 py-2 bg-blue-600 hover:bg-blue-700 text-white text-sm font-medium rounded-xl transition-colors shadow-sm disabled:opacity-70">
                                    {loading ? <Icon name="refresh-ccw" className="animate-spin" size={16} /> : <Icon name="save" size={16} />}
                                    Simpan
                                </button>
                            </div>
                        </div>
                    </div>
                );
            };

            const renderFuturisticConfirm = () => {
                if (!confirmDialog.isOpen) return null;

                return (
                    <div className="fixed inset-0 z-[100] flex items-center justify-center p-4">
                        <div className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm transition-opacity" onClick={() => setConfirmDialog({ ...confirmDialog, isOpen: false })}></div>
                        <div className="relative bg-white/95 dark:bg-slate-800/95 backdrop-blur-xl border border-slate-200/50 dark:border-slate-700/50 shadow-2xl rounded-3xl w-full max-w-sm p-8 overflow-hidden scale-in-center">
                            <div className={`absolute -top-20 -right-20 w-48 h-48 rounded-full blur-3xl opacity-20 ${confirmDialog.type === 'danger' ? 'bg-red-500' : 'bg-blue-500'}`}></div>
                            <div className={`absolute -bottom-20 -left-20 w-48 h-48 rounded-full blur-3xl opacity-10 ${confirmDialog.type === 'danger' ? 'bg-rose-500' : 'bg-indigo-500'}`}></div>

                            <div className="flex flex-col items-center text-center relative z-10">
                                <div className={`w-20 h-20 rounded-2xl flex items-center justify-center mb-5 shadow-inner backdrop-blur-md ${confirmDialog.type === 'danger' ? 'bg-red-50/80 text-red-500 shadow-red-200 border border-red-100' : 'bg-blue-50/80 text-blue-500 shadow-blue-200 border border-blue-100'}`}>
                                    <Icon name={confirmDialog.type === 'danger' ? 'alert-triangle' : 'help-circle'} size={36} />
                                </div>
                                <h3 className="text-xl font-bold text-slate-800 dark:text-slate-100 mb-2">{confirmDialog.title}</h3>
                                <p className="text-sm text-slate-500 mb-8 leading-relaxed px-2">{confirmDialog.message}</p>

                                <div className="flex w-full gap-3">
                                    <button
                                        onClick={() => setConfirmDialog({ ...confirmDialog, isOpen: false })}
                                        className="flex-1 py-3 rounded-xl text-sm font-bold text-slate-600 dark:text-slate-300 bg-slate-100/50 dark:bg-slate-700/50 hover:bg-slate-100 dark:hover:bg-slate-700 hover:text-slate-800 dark:hover:text-slate-100 transition-all border border-slate-200/60 dark:border-slate-600/50"
                                    >
                                        Batal
                                    </button>
                                    <button
                                        onClick={() => {
                                            if (confirmDialog.onConfirm) confirmDialog.onConfirm();
                                            setConfirmDialog({ ...confirmDialog, isOpen: false });
                                        }}
                                        className={`flex-1 py-3 rounded-xl text-sm font-bold text-white transition-all hover:scale-[1.03] active:scale-95 shadow-lg ${confirmDialog.type === 'danger' ? 'bg-gradient-to-br from-red-500 to-rose-600 shadow-red-500/30' : 'bg-gradient-to-br from-blue-500 to-indigo-600 shadow-blue-500/30'}`}
                                    >
                                        Konfirmasi
                                    </button>
                                </div>
                            </div>
                        </div>
                    </div>
                );
            };

            const renderDeveloperPromptModal = () => {
                if (!showDevPrompt) return null;

                return (
                    <div className="fixed inset-0 z-[100] flex items-center justify-center p-4">
                        <div className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm" onClick={() => setShowDevPrompt(false)}></div>
                        <div className="relative bg-white/95 dark:bg-slate-800/95 backdrop-blur-xl border border-slate-200/50 dark:border-slate-700/50 shadow-2xl rounded-3xl w-full max-w-sm p-8 overflow-hidden scale-in-center">
                            <div className="absolute -top-20 -right-20 w-48 h-48 rounded-full blur-3xl opacity-20 bg-indigo-500"></div>
                            <div className="flex flex-col items-center text-center relative z-10">
                                <div className="w-16 h-16 rounded-2xl flex items-center justify-center mb-4 bg-indigo-50 text-indigo-500 shadow-inner">
                                    <Icon name="alert-triangle" size={32} />
                                </div>
                                <h3 className="text-xl font-bold text-slate-800 dark:text-slate-100 mb-2">Akses Terbatas (BETA)</h3>
                                <p className="text-sm text-slate-500 mb-6 leading-relaxed">Menu {pendingTabAttempt === 'inventaris' ? 'Logistik & Inventaris' : pendingTabAttempt === 'ahli' ? 'Tenaga Ahli' : pendingTabAttempt === 'penugasan' ? 'Penugasan Tenaga Ahli' : 'Evaluasi & KPI'} masih dalam tahap uji coba (BETA) dan sedang dalam tahap pengembangan sementara hanya dapat diakses oleh Developer (Krisna).</p>

                                <input
                                    type="password"
                                    placeholder="Masukkan Password Developer"
                                    className="w-full p-3 rounded-xl border border-slate-300 dark:border-slate-600 bg-slate-50 dark:bg-slate-900/50 text-center text-sm font-semibold mb-2 focus:border-indigo-500 focus:ring-2 focus:ring-indigo-200 outline-none transition-all dark:text-slate-200"
                                    value={devPassword}
                                    onChange={(e) => {
                                        setDevPassword(e.target.value);
                                        setDevAuthError("");
                                    }}
                                    onKeyDown={(e) => { if (e.key === 'Enter') handleDevLogin(); }}
                                />
                                {devAuthError ? <p className="text-xs text-red-500 font-medium mb-4">{devAuthError}</p> : <div className="mb-4"></div>}

                                <div className="flex w-full gap-3">
                                    <button onClick={() => setShowDevPrompt(false)} className="flex-1 py-3 rounded-xl text-sm font-bold text-slate-600 dark:text-slate-300 bg-slate-100/50 dark:bg-slate-700/50 hover:bg-slate-100 dark:hover:bg-slate-700 transition-all border border-slate-200/60 dark:border-slate-600/50">Batal</button>
                                    <button onClick={handleDevLogin} className="flex-1 py-3 rounded-xl text-sm font-bold text-white bg-indigo-600 hover:bg-indigo-700 transition-all shadow-lg shadow-indigo-600/30">Akses</button>
                                </div>
                            </div>
                        </div>
                    </div>
                );
            };


            const renderDominoModal = () => {
                if (!dominoAnalysis) return null;
                const { project, delayDays, impactedEmployees } = dominoAnalysis;

                return (
                    <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
                        <div className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm" onClick={() => setDominoAnalysis(null)}></div>
                        <div className="relative bg-white dark:bg-slate-800 border border-slate-200/50 dark:border-slate-700/50 shadow-2xl rounded-3xl w-full max-w-2xl max-h-[90vh] flex flex-col overflow-hidden scale-in-center">
                            <div className="flex items-center justify-between p-5 border-b border-slate-100 dark:border-slate-700/50 bg-slate-50 dark:bg-slate-900/50">
                                <div className="flex items-center gap-3">
                                    <div className="w-10 h-10 rounded-xl bg-orange-100 flex items-center justify-center text-orange-600 shadow-sm">
                                        <Icon name="alert-triangle" size={20} />
                                    </div>
                                    <div>
                                        <h3 className="text-lg font-bold text-slate-800 dark:text-slate-100 leading-tight">Analisis Efek Domino 💥</h3>
                                        <p className="text-xs text-slate-500 font-medium">Prediksi tabrakan jadwal akibat keterlambatan proyek</p>
                                    </div>
                                </div>
                                <button onClick={() => setDominoAnalysis(null)} className="p-2 text-slate-400 hover:text-red-500 hover:bg-red-50 rounded-lg transition-colors"><Icon name="x" size={20} /></button>
                            </div>

                            <div className="p-6 overflow-y-auto custom-scrollbar flex-1 bg-white dark:bg-slate-800">
                                <div className="mb-6 bg-red-50 border border-red-100 p-4 rounded-xl">
                                    <p className="text-sm text-slate-700 leading-relaxed">
                                        Proyek <strong className="text-red-600">{project.name}</strong> saat ini terlambat <strong className="text-red-600 text-lg">{delayDays} hari</strong>.
                                    </p>
                                    <p className="text-xs text-slate-500 mt-1">
                                        Sistem memprediksi anggota sub-tim yang belum selesai di proyek ini akan membawa keterlambatan ini ke proyek mereka selanjutnya yang sedang berjalan.
                                    </p>
                                </div>

                                {impactedEmployees.length === 0 ? (
                                    <div className="flex flex-col items-center justify-center py-10 text-slate-400">
                                        <Icon name="check-circle-2" size={48} className="mb-3 text-emerald-400 opacity-50" />
                                        <h4 className="text-base font-bold text-slate-700 dark:text-slate-200">Aman Terkendali</h4>
                                        <p className="text-sm text-center max-w-sm mt-1">Tidak ditemukan efek domino. Anggota sub-tim tidak memiliki jadwal proyek lain di masa depan yang saling berdekatan.</p>
                                    </div>
                                ) : (
                                    <div className="space-y-6">
                                        {impactedEmployees.map((emp, i) => (
                                            <div key={i} className="bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 shadow-sm rounded-2xl overflow-hidden">
                                                <div className="bg-slate-50 dark:bg-slate-900/50 border-b border-slate-100 dark:border-slate-700 px-4 py-3 flex items-center gap-3">
                                                    <div className="w-8 h-8 rounded-full bg-blue-100 text-blue-600 flex items-center justify-center font-bold text-xs">{emp.name.substring(0, 2).toUpperCase()}</div>
                                                    <div>
                                                        <p className="text-sm font-bold text-slate-800 dark:text-slate-200">{emp.name}</p>
                                                        <p className="text-[10px] text-slate-500 font-medium">Beresiko membawa keterlambatan {delayDays} hari</p>
                                                    </div>
                                                </div>
                                                <div className="p-4 bg-slate-50/50 dark:bg-slate-900/20">
                                                    <p className="text-xs font-semibold text-slate-600 mb-3 uppercase tracking-wider">Proyek Masa Depan Yang Terancam Terlambat:</p>
                                                    <div className="space-y-3">
                                                        {emp.futureProjects.map((fp, j) => (
                                                            <div key={j} className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 p-3 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-xl shadow-sm relative overflow-hidden">
                                                                <div className="absolute left-0 top-0 bottom-0 w-1 bg-red-500"></div>
                                                                <div>
                                                                    <p className="text-sm font-bold text-slate-800 dark:text-slate-200">{fp.name}</p>
                                                                    <p className="text-xs text-slate-500 mt-0.5">Deadline: {formatDateIndo(fp.deadlineStr)} (Sisa <span className="font-semibold">{fp.remainingDays} hari</span>)</p>
                                                                </div>
                                                                <div className="bg-red-50 px-3 py-2 rounded-lg border border-red-100 text-right">
                                                                    <p className="text-[10px] text-red-600 font-bold uppercase tracking-wider mb-0.5">Prediksi Hasil</p>
                                                                    <p className="text-sm font-black text-red-700 flex items-center gap-1.5 justify-end">
                                                                        Terlambat {Math.abs(fp.clashDays)} Hari <Icon name="trending-down" size={14} />
                                                                    </p>
                                                                </div>
                                                            </div>
                                                        ))}
                                                    </div>
                                                </div>
                                            </div>
                                        ))}
                                    </div>
                                )}
                            </div>

                            <div className="p-5 border-t border-slate-100 dark:border-slate-700/50 bg-slate-50 dark:bg-slate-900/50 flex justify-end">
                                <button onClick={() => setDominoAnalysis(null)} className="px-5 py-2.5 rounded-xl text-sm font-bold text-slate-700 dark:text-slate-300 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 hover:bg-slate-50 dark:hover:bg-slate-700 transition-colors shadow-sm">Tutup Analisis</button>
                            </div>
                        </div>
                    </div>
                );
            };


            const renderKPIInfoModal = () => {
                if (!showKPIInfoModal) return null;

                return (
                    <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
                        <div className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm" onClick={() => setShowKPIInfoModal(false)}></div>
                        <div className="relative bg-white dark:bg-slate-800 border border-slate-200/50 dark:border-slate-700/50 shadow-2xl rounded-3xl w-full max-w-2xl max-h-[90vh] flex flex-col overflow-hidden scale-in-center">
                            <div className="flex items-center justify-between p-5 border-b border-slate-100 dark:border-slate-700/50 bg-slate-50 dark:bg-slate-900/50">
                                <div className="flex items-center gap-3">
                                    <div className="w-10 h-10 rounded-xl bg-blue-100 flex items-center justify-center text-blue-600 shadow-sm">
                                        <Icon name="info" size={20} />
                                    </div>
                                    <div>
                                        <h3 className="text-lg font-bold text-slate-800 dark:text-slate-100 leading-tight">Informasi Perhitungan KPI</h3>
                                        <p className="text-xs text-slate-500 font-medium">Aturan (Rules Point) Evaluasi Kinerja</p>
                                    </div>
                                </div>
                                <button onClick={() => setShowKPIInfoModal(false)} className="p-2 text-slate-400 hover:text-red-500 hover:bg-red-50 rounded-lg transition-colors"><Icon name="x" size={20} /></button>
                            </div>
                            <div className="p-6 overflow-y-auto custom-scrollbar flex-1 bg-white dark:bg-slate-800 space-y-4">
                                <div className="bg-slate-50 dark:bg-slate-900/50 p-4 rounded-2xl border border-slate-100 dark:border-slate-700/50">
                                    <h4 className="font-bold text-slate-800 dark:text-slate-200 mb-2 flex items-center gap-2"><Icon name="users" size={16} className="text-blue-500" /> Perhitungan Sebagai Staff / Anggota Tim</h4>
                                    <ul className="list-disc pl-5 text-sm text-slate-600 dark:text-slate-400 space-y-1.5">
                                        <li><strong>Skor Dasar:</strong> 100 Poin</li>
                                        <li><strong className="text-red-500">Penalti Keterlambatan:</strong> <strong className="font-bold">-5 Poin</strong> untuk setiap tugas yang melewati deadline dan belum mencapai 100%.</li>
                                        <li><strong className="text-emerald-500">Bonus Selesai:</strong> <strong className="font-bold">+15 Poin</strong> untuk setiap tugas yang selesai 100% tepat waktu atau lebih cepat.</li>
                                    </ul>
                                </div>
                                <div className="bg-amber-50/50 dark:bg-amber-900/10 p-4 rounded-2xl border border-slate-200 dark:border-amber-900/30">
                                    <h4 className="font-bold text-amber-900 dark:text-amber-400 mb-2 flex items-center gap-2"><Icon name="star" size={16} className="text-amber-500" /> Perhitungan Sebagai Team Leader</h4>
                                    <ul className="list-disc pl-5 text-sm text-slate-600 dark:text-slate-400 space-y-1.5">
                                        <li><strong>Kondisi 1: Saat memimpin proyek (Makro):</strong>
                                            <ul className="list-[circle] pl-5 mt-1 space-y-1">
                                                <li><strong className="text-red-500">Penalti Tanggung Jawab:</strong> <strong className="font-bold">-10 Poin</strong> jika proyek utama terlambat secara keseluruhan.</li>
                                                <li><strong className="text-emerald-500">Bonus Keberhasilan:</strong> <strong className="font-bold">+20 Poin</strong> jika proyek utama selesai 100% tepat waktu.</li>
                                            </ul>
                                        </li>
                                        <li className="pt-1"><strong>Kondisi 2: Saat bertugas sebagai sub-tim di proyek lain:</strong>
                                            <ul className="list-[circle] pl-5 mt-1 space-y-1">
                                                <li>Diberlakukan aturan Staff (<strong className="text-red-500">-5 Poin</strong> jika tugasnya terlambat, <strong className="text-emerald-500">+15 Poin</strong> jika tugasnya selesai).</li>
                                            </ul>
                                        </li>
                                    </ul>
                                </div>
                                <div className="bg-purple-50/50 dark:bg-purple-900/10 p-4 rounded-2xl border border-purple-100 dark:border-purple-900/30">
                                    <h4 className="font-bold text-purple-900 dark:text-purple-400 mb-2 flex items-center gap-2"><Icon name="zap" size={16} className="text-purple-500" /> Aturan Beban Kerja Berlebih (Overload)</h4>
                                    <p className="text-sm text-slate-600 dark:text-slate-400 mb-2">Berlaku untuk semua personil jika menangani <strong>lebih dari 4 proyek aktif</strong> secara bersamaan:</p>
                                    <ul className="list-disc pl-5 text-sm text-slate-600 dark:text-slate-400 space-y-1.5">
                                        <li><strong className="text-emerald-500">Bonus Produktivitas:</strong> Jika <strong>tidak ada satu pun</strong> proyek yang terlambat, mendapat tambahan <strong className="font-bold">+10 Poin per proyek ekstra</strong> (contoh: 5 proyek aktif = +10 poin, 6 proyek aktif = +20 poin dan seterusnya).</li>
                                        <li><strong className="text-red-500">Penalti Kewalahan:</strong> Jika <strong>terdapat proyek yang terlambat</strong> saat kondisi overload, akan langsung dijatuhi penalti tambahan <strong className="font-bold">-10 Poin</strong> (contoh: 5 proyek aktif, salah satunya terlambat = -10 poin, 6 proyek aktif, salah satunya terlambat = -20 poin dan seterusnya).</li>
                                    </ul>
                                </div>
                                <div className="bg-emerald-50/50 dark:bg-emerald-900/10 p-4 rounded-2xl border border-emerald-100 dark:border-emerald-900/30">
                                    <h4 className="font-bold text-emerald-900 dark:text-emerald-400 mb-2 flex items-center gap-2"><Icon name="bar-chart" size={16} className="text-emerald-500" /> Aturan Tie-Breaker</h4>
                                    <p className="text-sm text-slate-600 dark:text-slate-400 mb-2">Jika terdapat personil dengan Skor KPI yang <strong>sama persis</strong>, peringkat klasemen (leaderboard) akan ditentukan menggunakan prioritas berikut:</p>
                                    <ol className="list-decimal pl-5 text-sm text-slate-600 dark:text-slate-400 space-y-1.5 font-medium">
                                        <li><strong className="text-slate-800 dark:text-slate-200">Rating Bintang Manual:</strong> Personil dengan Bintang Rating Manual lebih tinggi akan diletakkan di atas.</li>
                                        <li><strong className="text-slate-800 dark:text-slate-200">Penyelesaian Terbanyak:</strong> Jika Bintang Rating sama, personil dengan jumlah proyek selesai 100% tanpa keterlambatan terbanyak akan menang.</li>
                                        <li><strong className="text-slate-800 dark:text-slate-200">Minim Keterlambatan:</strong> Jika jumlah penyelesaian sama, personil dengan jumlah proyek terlambat lebih sedikit berada di atas.</li>
                                        <li><strong className="text-slate-800 dark:text-slate-200">Rata-Rata Progress:</strong> Jika tingkat keterlambatan juga sama, personil dengan persentase progress lebih tinggi akan diprioritaskan.</li>
                                    </ol>
                                </div>
                                <div className="bg-slate-50 dark:bg-slate-900/50 p-3 rounded-xl border border-slate-200 dark:border-slate-700 flex items-start gap-2">
                                    <Icon name="info" size={16} className="text-slate-400 shrink-0 mt-0.5" />
                                    <p className="text-xs text-slate-500 dark:text-slate-400 italic">
                                        Skor akhir dibatasi maksimal 100 dan minimal 0. Skor dapat direvisi/di-override secara manual oleh Admin melalui opsi "Rating Manual" pada dashboard KPI.
                                    </p>
                                </div>
                            </div>
                            <div className="p-5 border-t border-slate-100 dark:border-slate-700/50 bg-slate-50 dark:bg-slate-900/50 flex justify-end">
                                <button onClick={() => setShowKPIInfoModal(false)} className="px-5 py-2.5 rounded-xl text-sm font-bold text-white bg-blue-600 hover:bg-blue-700 transition-colors shadow-lg shadow-blue-600/30">Mengerti</button>
                            </div>
                        </div>
                    </div>
                );
            };
            const handleExportExcel = (options) => {
                let filteredProjects = [...computedProjects].filter(p => p.computedStatus !== 'Done' && !p.notStarted);
                let filteredResources = [...calculatedResources];
                let title = "Laporan Eksekutif";

                if (options.section === 'ProyekSaja') title = "Laporan Eksekutif Status Proyek";
                if (options.section === 'PegawaiSaja') title = "Laporan Eksekutif Penugasan Personil";

                if (options.projectType !== 'Semua') {
                    title += ` (Tipe: ${options.projectType})`;
                    filteredProjects = filteredProjects.filter(p => {
                        const pType = (p.type || '').toLowerCase();
                        if (options.projectType === 'Perencanaan') return pType.includes('perencana');
                        if (options.projectType === 'Pengawasan') return pType === 'pengawasan' || pType === 'supervisi';
                        if (options.projectType === 'Manajemen Konstruksi') return pType.includes('manajemen konstruksi') || pType === 'mk';
                        return true;
                    });
                }
                
                filteredProjects.sort((a, b) => {
                    const dateA = a.spmk ? new Date(a.spmk).getTime() : 0;
                    const dateB = b.spmk ? new Date(b.spmk).getTime() : 0;
                    return dateB - dateA;
                });

                const aoa = [];
                const merges = [];
                const cardRanges = [];
                let rowIndex = 0;

                const addRow = (row, mergeToCol = -1) => {
                    aoa.push(row);
                    if (mergeToCol > 0) {
                        merges.push({ s: { r: rowIndex, c: 0 }, e: { r: rowIndex, c: mergeToCol } });
                    }
                    rowIndex++;
                };

                addRow([title], 7);
                addRow([`Tanggal Cetak: ${formatDateIndo(new Date().toISOString().split('T')[0])}`], 7);
                addRow([]);

                if (options.section !== 'PegawaiSaja') {
                    addRow(["BAGIAN A: LAPORAN STATUS PROYEK"], 7);
                    if (filteredProjects.length === 0) {
                        addRow(["Tidak ada proyek aktif untuk kriteria ini."], 7);
                    } else {
                        const bagianAStart = rowIndex;
                        const categoriesToRender = (options.projectType !== 'Semua' ? [options.projectType] : ['Perencanaan', 'Pengawasan', 'Manajemen Konstruksi']).map(kategoriTipe => {
                            return {
                                kategoriTipe,
                                proyeksInKategori: filteredProjects.filter(p => {
                                    const pType = (p.type || '').toLowerCase();
                                    if (kategoriTipe === 'Perencanaan') return pType.includes('perencana');
                                    if (kategoriTipe === 'Pengawasan') return pType === 'pengawasan' || pType === 'supervisi';
                                    if (kategoriTipe === 'Manajemen Konstruksi') return pType.includes('manajemen konstruksi') || pType === 'mk';
                                    return false;
                                })
                            };
                        }).filter(cat => cat.proyeksInKategori.length > 0);

                        if (options.projectType === 'Semua') {
                            const mappedIds = new Set(categoriesToRender.flatMap(g => g.proyeksInKategori.map(p => p.id)));
                            const others = filteredProjects.filter(p => !mappedIds.has(p.id));
                            if (others.length > 0) {
                                categoriesToRender.push({ kategoriTipe: 'Lainnya', proyeksInKategori: others });
                            }
                        }

                        categoriesToRender.forEach(cat => {
                            if (options.projectType === 'Semua') {
                                addRow([`Sub Bab: ${cat.kategoriTipe}`], 7);
                            }
                            
                            cat.proyeksInKategori.forEach(p => {
                                const isPengawasan = p.type?.toLowerCase().includes('pengawas') || p.type?.toLowerCase().includes('manajemen konstruksi');
                                
                                const startRow = rowIndex;
                                
                                addRow([`NAMA PROYEK: ${p.name || '-'}`], 7);
                                if (isPengawasan) {
                                    addRow([
                                        `Klien: ${p.client || '-'}`,
                                        `Tanggal SPMK: ${p.spmk ? formatDateIndo(p.spmk) : '-'}`,
                                        `Tenggat Kontrak: ${p.deadline ? formatDateIndo(p.deadline) : '-'}`,
                                        `Status Makro: ${p.computedStatus || '-'}`,
                                        "",
                                        "",
                                        "",
                                        ""
                                    ]);
                                } else {
                                    addRow([
                                        `Klien: ${p.client || '-'}`,
                                        `Team Leader: ${p.teamLeader || '-'}`,
                                        `Tanggal SPMK: ${p.spmk ? formatDateIndo(p.spmk) : '-'}`,
                                        `Tenggat Kontrak: ${p.deadline ? formatDateIndo(p.deadline) : '-'}`,
                                        `Status Makro: ${p.computedStatus || '-'}`,
                                        "",
                                        "",
                                        ""
                                    ]);
                                }
                                
                                addRow(["Rincian Penugasan & Target Sub-Tim"], 7);
                                
                                let sumStart, sumEnd;
                                
                                if (isPengawasan) {
                                    addRow(["Nama Personil", "Peran", "Man Month", "Deadline Spesifik", "Status Lapangan", "Gaji Total", "Gaji Bank", "Gaji Tunai"]);
                                    sumStart = rowIndex;
                                    if (!p.team || p.team.length === 0) {
                                        addRow(["Belum ada personil diplot.", "", "", "", "", "", "", ""], 7);
                                    } else {
                                        const sortedTeam = [...p.team].sort((a, b) => {
                                            const roleA = p.pengawasanDetails?.[a]?.role || 'Inspector';
                                            const roleB = p.pengawasanDetails?.[b]?.role || 'Inspector';
                                            return getLPSEHierarchyScore(roleA) - getLPSEHierarchyScore(roleB);
                                        });
                                        sortedTeam.forEach(member => {
                                            const details = p.pengawasanDetails?.[member] || {};
                                            addRow([
                                                member,
                                                details.role || 'Inspector',
                                                details.manMonth || '-',
                                                details.deadline ? formatDateIndo(details.deadline) : '-',
                                                details.statusTurun || 'Tidak Turun',
                                                "",
                                                "",
                                                ""
                                            ]);
                                        });
                                    }
                                } else {
                                    addRow(["Kategori Sub Tim", "Target Progress", "Tenggat Waktu Tim", "Status Target", "Personil Terploting", "Gaji Total", "Gaji Bank", "Gaji Tunai"]);
                                    sumStart = rowIndex;
                                    const cats = ['Arsitek', 'Struktur', 'MEP', 'QS', 'Tata Ruang', 'Lainnya', 'Surveyor'];
                                    let hasAny = false;
                                    cats.forEach(cat => {
                                        let catMembers = [];
                                        if (cat === 'Surveyor') {
                                            catMembers = p.surveyorTeam || [];
                                        } else {
                                            catMembers = (p.team || []).filter(m => {
                                                const isSurveyor = (p.surveyorTeam || []).includes(m);
                                                if (isSurveyor) return false;
                                                const r = calculatedResources.find(res => res.name === m);
                                                if (cat === 'Lainnya') {
                                                    const matchesOther = ['Arsitek', 'Struktur', 'MEP', 'QS', 'Tata Ruang'].some(other => getCategoryFromRole(r?.role) === other);
                                                    return !matchesOther;
                                                } else {
                                                    return getCategoryFromRole(r?.role) === cat;
                                                }
                                            });
                                        }
                                        if (catMembers.length > 0) {
                                            catMembers.sort((a, b) => {
                                                const resA = calculatedResources.find(r => r.name === a);
                                                const resB = calculatedResources.find(r => r.name === b);
                                                return getLPSEHierarchyScore(resA?.role) - getLPSEHierarchyScore(resB?.role);
                                            });
                                            hasAny = true;
                                            const details = p.categoryDetails?.[cat] || {};
                                            const progressVal = details.progress || 0;
                                            const statusTarget = getMicroStatus(progressVal, details.deadline);
                                            addRow([
                                                cat,
                                                `${progressVal}% Tercapai`,
                                                details.deadline ? formatDateIndo(details.deadline) : '-',
                                                statusTarget,
                                                catMembers.join(', '),
                                                "",
                                                "",
                                                ""
                                            ]);
                                        }
                                    });
                                    if (!hasAny) {
                                        addRow(["Belum ada personil diplot.", "", "", "", "", "", "", ""], 7);
                                    }
                                }
                                sumEnd = rowIndex - 1;
                                
                                const totalProyekRow = rowIndex;
                                addRow(["", "", "", "", "TOTAL PROYEK", { t: 'n', f: `SUBTOTAL(9,F${sumStart + 1}:F${sumEnd + 1})` }, { t: 'n', f: `SUBTOTAL(9,G${sumStart + 1}:G${sumEnd + 1})` }, { t: 'n', f: `SUBTOTAL(9,H${sumStart + 1}:H${sumEnd + 1})` }]);
                                
                                cardRanges.push({ s: { r: startRow, c: 0 }, e: { r: rowIndex - 1, c: 7 } });
                                
                                addRow([]);
                                addRow([]);
                            });
                        });
                        
                        const bagianAEnd = rowIndex - 1;
                        addRow([]);
                        const totalStartRow = rowIndex;
                        addRow(["", "", "", "", "TOTAL KESELURUHAN PROYEK", "Gaji Total", "Gaji Bank", "Gaji Tunai"]);
                        addRow(["", "", "", "", "", { t: 'n', f: `SUBTOTAL(9,F${bagianAStart + 1}:F${bagianAEnd + 1})` }, { t: 'n', f: `SUBTOTAL(9,G${bagianAStart + 1}:G${bagianAEnd + 1})` }, { t: 'n', f: `SUBTOTAL(9,H${bagianAStart + 1}:H${bagianAEnd + 1})` }]);
                        merges.push({ s: { r: totalStartRow, c: 4 }, e: { r: totalStartRow + 1, c: 4 } });
                        cardRanges.push({ s: { r: totalStartRow, c: 4 }, e: { r: totalStartRow + 1, c: 7 } });
                    }
                    addRow([]);
                }

                if (options.section !== 'ProyekSaja') {
                    addRow(["BAGIAN B: LAPORAN RINCIAN PENUGASAN PEGAWAI"], 7);
                    if (filteredResources.length === 0) {
                        addRow(["Tidak ada personil untuk kriteria ini."], 7);
                    } else {
                        const bagianBStart = rowIndex;
                        const resourcesBySubTeam = {};
                        filteredResources.forEach(res => {
                            const activeProjectsForRes = filteredProjects.filter(p => (p.team?.includes(res.name) || p.teamLeader === res.name || p.surveyorTeam?.includes(res.name)) && !p.notStarted);
                            if (options.projectType !== 'Semua' && activeProjectsForRes.length === 0) return;
                            
                            const subTeam = getCategoryFromRole(res.role);
                            if (!resourcesBySubTeam[subTeam]) resourcesBySubTeam[subTeam] = [];
                            resourcesBySubTeam[subTeam].push({ res, activeProjectsForRes });
                        });

                        const subTeamKeys = Object.keys(resourcesBySubTeam).sort();
                        if (subTeamKeys.length === 0) {
                            addRow(["Tidak ada personil yang terlibat di kriteria ini."], 7);
                        } else {
                            subTeamKeys.forEach(subTeam => {
                                addRow([`Sub Bab: Tim ${subTeam}`], 7);
                                const sortedResources = [...resourcesBySubTeam[subTeam]].sort((a, b) => getLPSEHierarchyScore(a.res.role) - getLPSEHierarchyScore(b.res.role) || a.res.name.localeCompare(b.res.name));
                                sortedResources.forEach(({ res, activeProjectsForRes }) => {
                                    const startRow = rowIndex;
                                    addRow([`Nama: ${res.name}`, `Peran: ${res.role}`, `Jumlah Proyek: ${activeProjectsForRes.length} Proyek Aktif`, "", "", "", "", ""]);
                                    merges.push({ s: { r: rowIndex - 1, c: 2 }, e: { r: rowIndex - 1, c: 4 } });
                                    
                                    addRow(["Nama Proyek", "Peran/Tim Lapangan", "Man Month", "Deadline Spesifik Tugas", "Status Lapangan", "Gaji Total", "Gaji Bank", "Gaji Tunai"]);
                                    
                                    if (activeProjectsForRes.length === 0) {
                                        addRow(["Sedang tidak memegang proyek aktif (Available).", "", "", "", "", "", "", ""], 7);
                                    } else {
                                        const categories = ['Perencanaan', 'Pengawasan', 'Manajemen Konstruksi'];
                                        const renderRowData = (p) => {
                                            const isPengawasan = p.type?.toLowerCase().includes('pengawas') || p.type?.toLowerCase().includes('manajemen konstruksi');
                                            let deadlineStr = '-';
                                            let roleStr = '-';
                                            let statusLapangan = '-';
                                            let manMonthStr = '-';

                                            if (p.teamLeader === res.name) {
                                                roleStr = 'Team Leader';
                                                if (isPengawasan) {
                                                    deadlineStr = p.deadline ? formatDateIndo(p.deadline) : '-';
                                                    const details = p.pengawasanDetails?.[res.name] || {};
                                                    statusLapangan = details.statusTurun || 'Tidak Turun';
                                                    manMonthStr = details.manMonth || '-';
                                                } else {
                                                    statusLapangan = p.computedStatus || '-';
                                                    const cats = ['Arsitek', 'QS', 'Struktur', 'MEP', 'Tata Ruang', 'Surveyor', 'Lainnya'];
                                                    let minDate = null;
                                                    let maxDate = null;
                                                    cats.forEach(cat => {
                                                        const d = p.categoryDetails?.[cat];
                                                        if (d) {
                                                            const startDateSource = d.startDate ? d.startDate : p.spmk;
                                                            if (startDateSource) {
                                                                const sd = new Date(startDateSource);
                                                                if (!minDate || sd < minDate) minDate = sd;
                                                            }
                                                            if (d.deadline) {
                                                                const ed = new Date(d.deadline);
                                                                if (!maxDate || ed > maxDate) maxDate = ed;
                                                            }
                                                        }
                                                    });
                                                    if (minDate && maxDate) {
                                                        minDate.setHours(0, 0, 0, 0);
                                                        maxDate.setHours(0, 0, 0, 0);
                                                        const diffDays = Math.ceil((maxDate.getTime() - minDate.getTime()) / (1000 * 60 * 60 * 24));
                                                        if (diffDays > 0) {
                                                            manMonthStr = parseFloat((diffDays / 30).toFixed(1)).toString();
                                                        }
                                                    }
                                                    deadlineStr = maxDate ? formatDateIndo(maxDate.toISOString().split('T')[0]) : (p.deadline ? formatDateIndo(p.deadline) : '-');
                                                }
                                            } else if (isPengawasan) {
                                                const details = p.pengawasanDetails?.[res.name] || {};
                                                deadlineStr = details.deadline ? formatDateIndo(details.deadline) : '-';
                                                roleStr = details.role || 'Inspector';
                                                statusLapangan = details.statusTurun || 'Tidak Turun';
                                                manMonthStr = details.manMonth || '-';
                                            } else {
                                                const effectiveCat = getEffectiveEmpCategory(p, res.name, res.role);
                                                const details = p.categoryDetails?.[effectiveCat] || {};
                                                deadlineStr = details.deadline ? formatDateIndo(details.deadline) : '-';
                                                roleStr = effectiveCat === 'Surveyor' ? 'Tim Surveyor' : effectiveCat;

                                                const startDateSource = details.startDate ? details.startDate : p.spmk;
                                                if (startDateSource && details.deadline) {
                                                    const startD = new Date(startDateSource);
                                                    const endD = new Date(details.deadline);
                                                    startD.setHours(0, 0, 0, 0);
                                                    endD.setHours(0, 0, 0, 0);
                                                    const diffDays = Math.ceil((endD.getTime() - startD.getTime()) / (1000 * 60 * 60 * 24));
                                                    if (diffDays > 0) {
                                                        manMonthStr = parseFloat((diffDays / 30).toFixed(1)).toString();
                                                    }
                                                }
                                                statusLapangan = getMicroStatus(details.progress || 0, details.deadline);
                                            }

                                            return [p.name, roleStr, manMonthStr, deadlineStr, statusLapangan, "", "", ""];
                                        };

                                        const dataStartRow = rowIndex;
                                        
                                        if (options.projectType === 'Semua') {
                                            const grouped = categories.map(kategoriTipe => {
                                                return {
                                                    kategoriTipe,
                                                    proyeks: activeProjectsForRes.filter(p => {
                                                        const pType = (p.type || '').toLowerCase();
                                                        if (kategoriTipe === 'Perencanaan') return pType.includes('perencana');
                                                        if (kategoriTipe === 'Pengawasan') return pType === 'pengawasan' || pType === 'supervisi';
                                                        if (kategoriTipe === 'Manajemen Konstruksi') return pType.includes('manajemen konstruksi') || pType === 'mk';
                                                        return false;
                                                    })
                                                };
                                            }).filter(cat => cat.proyeks.length > 0);
                                            
                                            const mappedIds = new Set(grouped.flatMap(g => g.proyeks.map(p => p.id)));
                                            const others = activeProjectsForRes.filter(p => !mappedIds.has(p.id));
                                            if (others.length > 0) {
                                                grouped.push({ kategoriTipe: 'Lainnya', proyeks: others });
                                            }

                                            grouped.forEach(cat => {
                                                addRow([cat.kategoriTipe, "", "", "", "", "", "", ""], 7);
                                                cat.proyeks.forEach(p => {
                                                    addRow(renderRowData(p));
                                                });
                                            });
                                        } else {
                                            activeProjectsForRes.forEach(p => {
                                                addRow(renderRowData(p));
                                            });
                                        }
                                        
                                        const dataEndRow = rowIndex - 1;
                                        const totalPegawaiRow = rowIndex;
                                        addRow(["", "", "", "", "TOTAL PEGAWAI", { t: 'n', f: `SUBTOTAL(9,F${dataStartRow + 1}:F${dataEndRow + 1})` }, { t: 'n', f: `SUBTOTAL(9,G${dataStartRow + 1}:G${dataEndRow + 1})` }, { t: 'n', f: `SUBTOTAL(9,H${dataStartRow + 1}:H${dataEndRow + 1})` }]);
                                    }
                                    cardRanges.push({ s: { r: startRow, c: 0 }, e: { r: rowIndex - 1, c: 7 } });
                                    
                                    addRow([]);
                                    addRow([]);
                                });
                            });
                            
                            const bagianBEnd = rowIndex - 1;
                            addRow([]);
                            const totalStartRow = rowIndex;
                            addRow(["", "", "", "", "TOTAL KESELURUHAN PEGAWAI", "Gaji Total", "Gaji Bank", "Gaji Tunai"]);
                            addRow(["", "", "", "", "", { t: 'n', f: `SUBTOTAL(9,F${bagianBStart + 1}:F${bagianBEnd + 1})` }, { t: 'n', f: `SUBTOTAL(9,G${bagianBStart + 1}:G${bagianBEnd + 1})` }, { t: 'n', f: `SUBTOTAL(9,H${bagianBStart + 1}:H${bagianBEnd + 1})` }]);
                            merges.push({ s: { r: totalStartRow, c: 4 }, e: { r: totalStartRow + 1, c: 4 } });
                            cardRanges.push({ s: { r: totalStartRow, c: 4 }, e: { r: totalStartRow + 1, c: 7 } });
                        }
                    }
                }

                const ws = XLSX.utils.aoa_to_sheet(aoa);
                ws['!merges'] = merges;
                
                // Add Styling for Borders and Bold
                for (let R = 0; R < rowIndex; ++R) {
                    let isEmptyRow = true;
                    for (let C = 0; C < 8; ++C) {
                        const cellRef = XLSX.utils.encode_cell({r: R, c: C});
                        if (ws[cellRef] && ((ws[cellRef].v !== undefined && ws[cellRef].v !== "") || ws[cellRef].f !== undefined)) {
                            isEmptyRow = false;
                            break;
                        }
                    }

                    if (!isEmptyRow) {
                        // Find if this row is part of a card
                        let card = cardRanges.find(c => R >= c.s.r && R <= c.e.r);
                        
                        for (let C = 0; C < 8; ++C) {
                            const cellRef = XLSX.utils.encode_cell({r: R, c: C});
                            if (!ws[cellRef]) ws[cellRef] = { t: 's', v: '' };
                            
                            let isBold = false;
                            let align = { vertical: "center", wrapText: true };
                            let fill = undefined;
                            
                            const val = ws[cellRef].v;
                            if (typeof val === 'string') {
                                if (val.startsWith('NAMA PROYEK:') || val.startsWith('BAGIAN') || val.startsWith('Sub Bab:') || val.includes('Laporan Eksekutif') || val === 'Rincian Penugasan & Target Sub-Tim') {
                                    isBold = true;
                                }
                                if (val.startsWith('NAMA PROYEK:') || val.startsWith('Nama:') || val.startsWith('Peran:') || val.startsWith('Jumlah Proyek:')) {
                                    fill = { fgColor: { rgb: "F1F5F9" } }; // Light gray background for card headers
                                }
                                
                                const catSeparators = ['Perencanaan', 'Pengawasan', 'Manajemen Konstruksi', 'Lainnya'];
                                if (catSeparators.includes(val)) {
                                    const nextCellRef = XLSX.utils.encode_cell({r: R, c: 1});
                                    if (!ws[nextCellRef] || ws[nextCellRef].v === "") {
                                        isBold = true;
                                        align.horizontal = "center";
                                        fill = { fgColor: { rgb: "F1F5F9" } };
                                    }
                                }

                                const headers = ["Nama Personil", "Peran", "Man Month", "Deadline Spesifik", "Status Lapangan", "Gaji Total", "Gaji Bank", "Gaji Tunai", "Kategori Sub Tim", "Target Progress", "Tenggat Waktu Tim", "Status Target", "Personil Terploting", "Nama Proyek", "Peran/Tim Lapangan", "Deadline Spesifik Tugas"];
                                if (headers.includes(val)) {
                                    isBold = true;
                                    align.horizontal = "center";
                                }
                                if (val.startsWith("Klien:") || val.startsWith("Team Leader:") || val.startsWith("Tanggal SPMK:") || val.startsWith("Tenggat Kontrak:") || val.startsWith("Status Makro:") || val.startsWith("Nama:") || val.startsWith("Peran:") || val.startsWith("Jumlah Proyek:")) {
                                    isBold = true;
                                }
                            }

                            // Determine borders based on card position
                            let bTop, bBottom, bLeft, bRight;
                            if (card) {
                                if (C >= card.s.c && C <= card.e.c) {
                                    bTop = 'thin'; bBottom = 'thin'; bLeft = 'thin'; bRight = 'thin';
                                    if (R === card.s.r) bTop = 'medium';
                                    if (R === card.e.r) bBottom = 'medium';
                                    if (C === card.s.c) bLeft = 'medium';
                                    if (C === card.e.c) bRight = 'medium';
                                }
                            } else {
                                bTop = 'medium'; bBottom = 'medium'; bLeft = 'medium'; bRight = 'medium';
                            }

                            ws[cellRef].s = {
                                font: { bold: isBold },
                                fill: fill,
                                alignment: align
                            };
                            
                            if (bTop) {
                                ws[cellRef].s.border = {
                                    top: { style: bTop, color: { rgb: "000000" } },
                                    bottom: { style: bBottom, color: { rgb: "000000" } },
                                    left: { style: bLeft, color: { rgb: "000000" } },
                                    right: { style: bRight, color: { rgb: "000000" } }
                                };
                            }
                        }
                    }
                }

                // Set column widths for better visibility
                ws['!cols'] = [
                    {wch: 35}, // Nama Proyek
                    {wch: 20}, // Peran
                    {wch: 15}, // Man Month
                    {wch: 20}, // Deadline
                    {wch: 15}, // Status Lapangan
                    {wch: 15}, // Gaji Total
                    {wch: 15}, // Gaji Bank
                    {wch: 15}  // Gaji Tunai
                ];

                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, "Laporan");
                XLSX.writeFile(wb, `Laporan_${new Date().getTime()}.xlsx`);
            };

            const renderPrintExecutiveReport = () => {
                if (!printData) return null;

                let title = "Laporan Eksekutif Status Proyek & Penugasan Personil";
                let filteredProjects = [...computedProjects].filter(p => p.computedStatus !== 'Done' && !p.notStarted);
                let filteredResources = [...calculatedResources];

                if (printData.type === 'project') {
                    filteredProjects = filteredProjects.filter(p => p.id === printData.id);
                    title = `Laporan Rincian Proyek: ${filteredProjects[0]?.name}`;
                    filteredResources = filteredResources.filter(r => filteredProjects[0]?.team?.includes(r.name));
                } else if (printData.type === 'personnel') {
                    filteredResources = filteredResources.filter(r => r.name === printData.name);
                    title = `Laporan Rekam Penugasan Personil: ${filteredResources[0]?.name}`;
                    filteredProjects = filteredProjects.filter(p => (p.team?.includes(filteredResources[0]?.name) || p.teamLeader === filteredResources[0]?.name) && !p.notStarted);
                } else if (printData.type === 'expert_assignment') {
                    title = "Laporan Eksekutif Penugasan Tenaga Ahli";
                } else if (printData.type === 'custom') {
                    if (printData.options.section === 'ProyekSaja') title = "Laporan Eksekutif Status Proyek";
                    if (printData.options.section === 'PegawaiSaja') title = "Laporan Eksekutif Penugasan Personil";

                    if (printData.options.projectType !== 'Semua') {
                        title += ` (Tipe: ${printData.options.projectType})`;
                        filteredProjects = filteredProjects.filter(p => {
                            const pType = (p.type || '').toLowerCase();
                            if (printData.options.projectType === 'Perencanaan') return pType.includes('perencana');
                            if (printData.options.projectType === 'Pengawasan') return pType === 'pengawasan' || pType === 'supervisi';
                            if (printData.options.projectType === 'Manajemen Konstruksi') return pType.includes('manajemen konstruksi') || pType === 'mk';
                            return true;
                        });
                    }
                }

                // Sort proyek berdasarkan deadline
                filteredProjects.sort((a, b) => {
                    const dateA = a.deadline ? new Date(a.deadline) : new Date(8640000000000000);
                    const dateB = b.deadline ? new Date(b.deadline) : new Date(8640000000000000);
                    return dateA - dateB;
                });

                const normalizeProjectType = (typeStr) => {
                    if (!typeStr) return '-';
                    const lower = typeStr.toLowerCase();
                    if (lower.includes('perencana')) return 'Perencanaan';
                    if (lower.includes('pengawas') || lower === 'supervisi') return 'Pengawasan';
                    if (lower.includes('manajemen konstruksi') || lower === 'mk') return 'Manajemen Konstruksi';
                    return typeStr;
                };

                const renderProjectPrintCard = (p) => {
                    const isPengawasan = p.type?.toLowerCase().includes('pengawas') || p.type?.toLowerCase().includes('manajemen konstruksi');
                    return (
                        <div key={p.id} className="border-2 border-slate-300 dark:border-slate-700 p-4" style={{ pageBreakInside: 'avoid' }}>
                            <div className="flex justify-between items-start border-b border-gray-300 pb-2 mb-3">
                                <div>
                                    <h3 className="text-base font-bold uppercase">{p.name}</h3>
                                    <p className="text-xs uppercase tracking-wide mt-0.5">
                                        {normalizeProjectType(p.type)} {' | '}Klien: {p.client || '-'}
                                    </p>
                                    {p.type?.toLowerCase().includes('perencana') && p.divisiKontrol && (
                                        <p className="text-xs font-bold mt-1 text-gray-800">Divisi Kontrol: {p.divisiKontrol}</p>
                                    )}
                                    {!isPengawasan && <p className="text-xs font-bold mt-1 text-gray-800">Team Leader: {p.teamLeader || 'Belum Ditentukan'}</p>}
                                </div>
                                <div className="text-right shrink-0 ml-4">
                                    <p className="text-xs font-bold border border-slate-300 dark:border-slate-700 px-2 py-1 inline-block uppercase whitespace-nowrap">Status Makro: {p.computedStatus}</p>

                                    {(isPengawasan || p.type?.toLowerCase().includes('perencana')) && <p className="text-xs mt-1 whitespace-nowrap font-semibold text-gray-800">Tanggal SPMK: {p.spmk ? formatDateIndo(p.spmk) : '-'}</p>}
                                    <p className="text-xs mt-1 whitespace-nowrap">Tenggat Kontrak: {p.deadline ? formatDateIndo(p.deadline) : '-'}</p>
                                </div>
                            </div>

                            <div>
                                <h4 className="text-[11px] font-bold uppercase mb-2">Rincian Penugasan & Target Sub-Tim</h4>
                                {isPengawasan ? (
                                    <table className="w-full text-xs border-collapse border border-slate-300 dark:border-slate-700 mb-1">
                                        <thead>
                                            <tr className="bg-gray-100">
                                                <th className="border border-slate-300 dark:border-slate-700 p-1.5 text-left w-[30%]">Nama Personil</th>
                                                <th className="border border-slate-300 dark:border-slate-700 p-1.5 text-left w-[20%]">Peran</th>
                                                <th className="border border-slate-300 dark:border-slate-700 p-1.5 text-left w-[15%]">Man Month</th>
                                                <th className="border border-slate-300 dark:border-slate-700 p-1.5 text-left w-[20%]">Deadline Spesifik</th>
                                                <th className="border border-slate-300 dark:border-slate-700 p-1.5 text-left w-[15%]">Status Lapangan</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {[...(p.team || [])].sort((a, b) => {
                                                const roleA = p.pengawasanDetails?.[a]?.role || 'Inspector';
                                                const roleB = p.pengawasanDetails?.[b]?.role || 'Inspector';
                                                return getLPSEHierarchyScore(roleA) - getLPSEHierarchyScore(roleB);
                                            }).map(member => {
                                                const details = p.pengawasanDetails?.[member] || {};
                                                return (
                                                    <tr key={member}>
                                                        <td className="border border-slate-300 dark:border-slate-700 p-1.5 font-bold">{member}</td>
                                                        <td className="border border-slate-300 dark:border-slate-700 p-1.5">{details.role || 'Inspector'}</td>
                                                        <td className="border border-slate-300 dark:border-slate-700 p-1.5 text-center">{details.manMonth || '-'}</td>
                                                        <td className="border border-slate-300 dark:border-slate-700 p-1.5">{details.deadline ? formatDateIndo(details.deadline) : '-'}</td>
                                                        <td className="border border-slate-300 dark:border-slate-700 p-1.5">{details.statusTurun || 'Tidak Turun'}</td>
                                                    </tr>
                                                )
                                            })}
                                            {(!p.team || p.team.length === 0) && (
                                                <tr><td colSpan="5" className="border border-slate-300 dark:border-slate-700 p-1.5 text-center italic">Belum ada personil diplot.</td></tr>
                                            )}
                                        </tbody>
                                    </table>
                                ) : (
                                    <table className="w-full text-xs border-collapse border border-slate-300 dark:border-slate-700 mb-1">
                                        <thead>
                                            <tr className="bg-gray-100">
                                                <th className="border border-slate-300 dark:border-slate-700 p-1.5 text-left w-[20%]">Kategori Sub Tim</th>
                                                <th className="border border-slate-300 dark:border-slate-700 p-1.5 text-left w-[15%]">Target Progress</th>
                                                <th className="border border-slate-300 dark:border-slate-700 p-1.5 text-left w-[20%]">Tenggat Waktu Tim</th>
                                                <th className="border border-slate-300 dark:border-slate-700 p-1.5 text-left w-[20%]">Status Target</th>
                                                <th className="border border-slate-300 dark:border-slate-700 p-1.5 text-left w-[25%]">Personil Terploting</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {['Arsitek', 'Struktur', 'MEP', 'QS', 'Tata Ruang', 'Lainnya', 'Surveyor'].map(cat => {
                                                let catMembers = [];
                                                if (cat === 'Surveyor') {
                                                    catMembers = p.surveyorTeam || [];
                                                } else {
                                                    catMembers = (p.team || []).filter(m => {
                                                        const isSurveyor = (p.surveyorTeam || []).includes(m);
                                                        if (isSurveyor) return false;
                                                        const r = resources.find(res => res.name === m);

                                                        if (cat === 'Lainnya') {
                                                            const matchesOther = ['Arsitek', 'Struktur', 'MEP', 'QS', 'Tata Ruang'].some(other => getCategoryFromRole(r?.role) === other);
                                                            return !matchesOther;
                                                        } else {
                                                            return getCategoryFromRole(r?.role) === cat;
                                                        }
                                                    });
                                                }

                                                if (catMembers.length === 0) return null;

                                                catMembers.sort((a, b) => {
                                                    const resA = calculatedResources.find(r => r.name === a);
                                                    const resB = calculatedResources.find(r => r.name === b);
                                                    return getLPSEHierarchyScore(resA?.role) - getLPSEHierarchyScore(resB?.role);
                                                });

                                                const details = p.categoryDetails?.[cat] || {};
                                                const progressVal = details.progress || 0;
                                                const statusTarget = getMicroStatus(progressVal, details.deadline);

                                                return (
                                                    <tr key={cat}>
                                                        <td className="border border-slate-300 dark:border-slate-700 p-1.5 font-bold">{cat}</td>
                                                        <td className="border border-slate-300 dark:border-slate-700 p-1.5">
                                                            {progressVal}% Tercapai
                                                        </td>
                                                        <td className="border border-slate-300 dark:border-slate-700 p-1.5">{details.deadline ? formatDateIndo(details.deadline) : '-'}</td>
                                                        <td className="border border-slate-300 dark:border-slate-700 p-1.5 font-semibold">{statusTarget}</td>
                                                        <td className="border border-slate-300 dark:border-slate-700 p-1.5">{catMembers.join(', ')}</td>
                                                    </tr>
                                                )
                                            })}
                                            {(!p.team || p.team.length === 0) && (
                                                <tr><td colSpan="5" className="border border-slate-300 dark:border-slate-700 p-1.5 text-center italic">Belum ada personil diplot.</td></tr>
                                            )}
                                        </tbody>
                                    </table>
                                )}
                            </div>
                        </div>
                    );
                };

                const renderPrintTimeSchedule = (p) => {
                    if (!p.type?.toLowerCase().includes('perencana')) return null;

                    const phases = [];
                    const possibleCats = ['Arsitek', 'Struktur', 'MEP', 'Tata Ruang', 'QS', 'Surveyor', 'Lainnya'];

                    let globalMinDate = null;
                    let globalMaxDate = null;

                    possibleCats.forEach(cat => {
                        const detail = p.categoryDetails?.[cat];
                        const hasData = detail && (detail.startDate || detail.deadline || (detail.progress !== undefined && detail.progress !== ""));
                        const hasSpmkData = p.spmk && detail && (detail.deadline || (detail.progress !== undefined && detail.progress !== ""));

                        if (hasData || hasSpmkData) {
                            const rawStartDate = detail.startDate ? detail.startDate : p.spmk;
                            let sd = rawStartDate ? new Date(rawStartDate) : null;
                            let ed = detail.deadline ? new Date(detail.deadline) : null;

                            if (sd && (!globalMinDate || sd < globalMinDate)) globalMinDate = new Date(sd);
                            if (ed && (!globalMaxDate || ed > globalMaxDate)) globalMaxDate = new Date(ed);
                            if (sd && !globalMaxDate) globalMaxDate = new Date(sd);
                            if (ed && !globalMinDate) globalMinDate = new Date(ed);

                            phases.push({
                                id: cat,
                                name: cat,
                                startDate: sd,
                                endDate: ed,
                                progress: detail.progress || 0,
                                isSubTask: false
                            });

                            if (detail.tasks && detail.tasks.length > 0) {
                                detail.tasks.forEach((task, tIdx) => {
                                    const isTaskCompleted = (detail.completedTasks || []).includes(task);
                                    phases.push({
                                        id: `${cat}-task-${tIdx}`,
                                        name: task,
                                        startDate: null,
                                        endDate: null,
                                        progress: isTaskCompleted ? 100 : 0,
                                        isSubTask: true
                                    });
                                });
                            }
                        }
                    });

                    if (phases.length === 0) return null;

                    if (!globalMinDate) globalMinDate = new Date();
                    if (!globalMaxDate) {
                        globalMaxDate = new Date();
                        globalMaxDate.setMonth(globalMaxDate.getMonth() + 3);
                    }

                    const chartStart = new Date(globalMinDate);
                    chartStart.setDate(1);

                    const chartEnd = new Date(globalMaxDate);
                    chartEnd.setMonth(chartEnd.getMonth() + 2);
                    chartEnd.setDate(0);

                    const timeUnits = [];
                    let curr = new Date(chartStart);
                    if (scheduleZoom === 'month') {
                        while (curr <= chartEnd) {
                            timeUnits.push({ date: new Date(curr), label: curr.toLocaleDateString('id-ID', { month: 'short', year: '2-digit' }).toUpperCase(), monthLabel: curr.toLocaleDateString('id-ID', { month: 'short', year: '2-digit' }).toUpperCase() });
                            curr.setMonth(curr.getMonth() + 1);
                        }
                    } else {
                        let currentMonth = -1;
                        let weekCounter = 1;
                        while (curr <= chartEnd) {
                            const monthStr = curr.toLocaleDateString('id-ID', { month: 'long', year: 'numeric' }).toUpperCase();
                            if (curr.getMonth() !== currentMonth) {
                                currentMonth = curr.getMonth();
                                weekCounter = 1;
                            }
                            timeUnits.push({ date: new Date(curr), label: `MG ${weekCounter}`, monthLabel: monthStr });
                            curr.setDate(curr.getDate() + 7);
                            weekCounter++;
                        }
                    }
                    const totalDays = Math.round((chartEnd - chartStart) / (1000 * 60 * 60 * 24)) + 1;

                    return (
                        <div className="print-landscape mt-10">
                            <h2 className="text-lg font-bold uppercase border-b border-gray-400 pb-2 mb-4">Laporan Time Schedule: {p.name}</h2>
                            <table className="w-full text-[10px] border-collapse border border-slate-300 dark:border-slate-700 table-fixed bg-white relative">
                                <thead>
                                    {scheduleZoom === 'week' ? (
                                        <>
                                            <tr className="bg-gray-200">
                                                <th rowSpan={2} className="border border-slate-300 dark:border-slate-700 p-1 w-[150px] text-left">Tasks / Kegiatan</th>
                                                <th rowSpan={2} className="border border-slate-300 dark:border-slate-700 p-1 w-[70px] text-center">Durasi</th>
                                                <th rowSpan={2} className="border border-slate-300 dark:border-slate-700 p-1 w-[105px] text-center">Status Target</th>
                                                <th rowSpan={2} className="border border-slate-300 dark:border-slate-700 p-1 w-[40px] text-center">Prog.</th>
                                                {(() => {
                                                    const monthGroups = [];
                                                    timeUnits.forEach(u => {
                                                        const last = monthGroups[monthGroups.length - 1];
                                                        if (last && last.monthLabel === u.monthLabel) last.span++;
                                                        else monthGroups.push({ monthLabel: u.monthLabel, span: 1 });
                                                    });
                                                    return monthGroups.map((g, i) => (
                                                        <th key={`m-${i}`} colSpan={g.span} className="border border-slate-300 dark:border-slate-700 p-1 text-center font-bold bg-gray-300">
                                                            {g.monthLabel}
                                                        </th>
                                                    ));
                                                })()}
                                            </tr>
                                            <tr className="bg-gray-200">
                                                {timeUnits.map((u, i) => (
                                                    <th key={`w-${i}`} className="border border-slate-300 dark:border-slate-700 p-1 text-center font-bold text-[8px]">
                                                        {u.label}
                                                    </th>
                                                ))}
                                            </tr>
                                        </>
                                    ) : (
                                        <tr className="bg-gray-200">
                                            <th className="border border-slate-300 dark:border-slate-700 p-1 w-[150px] text-left">Tasks / Kegiatan</th>
                                            <th className="border border-slate-300 dark:border-slate-700 p-1 w-[70px] text-center">Durasi</th>
                                            <th className="border border-slate-300 dark:border-slate-700 p-1 w-[105px] text-center">Status Target</th>
                                            <th className="border border-slate-300 dark:border-slate-700 p-1 w-[40px] text-center">Prog.</th>
                                            {timeUnits.map((u, i) => (
                                                <th key={i} className="border border-slate-300 dark:border-slate-700 p-1 text-center font-bold">
                                                    {u.label}
                                                </th>
                                            ))}
                                        </tr>
                                    )}
                                </thead>
                                <tbody>
                                    {phases.map((phase) => {
                                        const hasBar = phase.startDate && phase.endDate;
                                        let leftPerc = 0;
                                        let widthPerc = 0;
                                        let manMonthStr = '-';

                                        if (hasBar) {
                                            let s = new Date(phase.startDate);
                                            let e = new Date(phase.endDate);

                                            // Calculate Duration based on actual dates
                                            const daysTotal = Math.round((e - s) / (1000 * 60 * 60 * 24));
                                            if (daysTotal >= 0) {
                                                manMonthStr = daysTotal + ' Hari';
                                            }

                                            // Bounding for chart bar
                                            if (s < chartStart) s = chartStart;
                                            if (e > chartEnd) e = chartEnd;

                                            const offsetDays = Math.round((s - chartStart) / (1000 * 60 * 60 * 24));
                                            const durationDays = Math.round((e - s) / (1000 * 60 * 60 * 24)) + 1;

                                            leftPerc = (offsetDays / totalDays) * 100;
                                            widthPerc = (durationDays / totalDays) * 100;
                                        }

                                        return (
                                            <tr key={phase.id} className={phase.isSubTask ? "bg-white" : "bg-gray-50 font-semibold"}>
                                                <td className={`border border-slate-300 dark:border-slate-700 p-1.5 truncate ${phase.isSubTask ? 'pl-4 text-gray-700' : ''}`}>
                                                    {phase.isSubTask ? `└ ${phase.name}` : phase.name}
                                                </td>
                                                <td className="border border-slate-300 dark:border-slate-700 p-1 text-center font-bold whitespace-nowrap">{manMonthStr}</td>
                                                <td className="border border-slate-300 dark:border-slate-700 p-1 text-center font-bold uppercase">{getMicroStatus(phase.progress, phase.endDate)}</td>
                                                <td className="border border-slate-300 dark:border-slate-700 p-1 text-center">{phase.progress}%</td>
                                                <td colSpan={timeUnits.length} className="border border-slate-300 dark:border-slate-700 relative p-0 h-6">
                                                    <div className="absolute inset-0 flex z-0">
                                                        {timeUnits.map((m, i) => (
                                                            <div key={i} className="flex-1 border-l border-gray-300 border-dashed first:border-l-0"></div>
                                                        ))}
                                                    </div>
                                                    {hasBar && (
                                                        <div
                                                            className="absolute top-[4px] bottom-[4px] bg-black border border-slate-300 dark:border-slate-700 z-10 flex items-center justify-between px-1 text-[7px] text-white overflow-hidden whitespace-nowrap"
                                                            style={{ left: `${leftPerc}%`, width: `${widthPerc}%`, printColorAdjust: 'exact' }}
                                                        >
                                                            <span>{formatDateIndo(new Date(phase.startDate).toISOString().split('T')[0])}</span>
                                                            <span>{formatDateIndo(new Date(phase.endDate).toISOString().split('T')[0])}</span>
                                                        </div>
                                                    )}
                                                </td>
                                            </tr>
                                        )
                                    })}
                                </tbody>
                            </table>
                            <div className="mt-4 text-xs italic text-gray-600">
                                Catatan: Bar berwarna solid hitam menunjukkan rentang waktu pelaksanaan tugas secara proporsional. <br />
                                <strong>Tanggal Mulai (awal garis bar) pada chart ini disesuaikan dengan Tanggal Mulai masing-masing sub-tim.</strong>
                            </div>
                        </div>
                    );
                };

                  return (
                    <div id="print-wrapper" className="fixed inset-0 z-[9999] bg-slate-200 overflow-y-auto w-full h-screen font-sans pb-20 print:static print:bg-white print:h-auto print:overflow-visible print:pb-0">
                        {/* Control Bar for Print Preview */}
                        <div className="sticky top-0 w-full bg-white shadow-md p-4 flex justify-between items-center z-[10000] print:hidden shrink-0 border-b border-slate-200">
                            <div className="flex items-center gap-3">
                                <div className="p-2 bg-indigo-100 rounded-lg text-indigo-600"><Icon name="printer" size={20} /></div>
                                <div>
                                    <h2 className="text-lg font-bold text-slate-800 leading-tight">Pratinjau Laporan</h2>
                                    <p className="text-xs text-slate-500 font-medium">Klik Simpan PDF jika tampilan sudah sesuai.</p>
                                </div>
                            </div>
                            <div className="flex gap-3">
                                <button onClick={() => setPrintData(null)} className="px-5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold rounded-xl transition-all shadow-sm text-sm flex items-center gap-2">Tutup Pratinjau</button>
                                <button onClick={() => window.print()} className="px-5 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-xl transition-all shadow-md flex items-center gap-2 text-sm">
                                    <Icon name="printer" size={16} /> Simpan PDF / Cetak
                                </button>
                            </div>
                        </div>

                        {/* Page Wrapper */}
                        <div className="bg-white shadow-xl mt-8 mx-auto print:mt-0 print:mx-0 print:shadow-none relative h-fit" style={{ width: '100%', maxWidth: '297mm', minHeight: '210mm', backgroundColor: '#fff' }}>
                            <style type="text/css">
                                {`
                                @media print {
                                    #main-ui-wrapper { display: none !important; }
                                    #print-wrapper { position: static !important; }
                                    @page { margin: 0; }
                                    body { 
                                        -webkit-print-color-adjust: exact; 
                                        print-color-adjust: exact; 
                                        padding: 0; margin: 0; 
                                        background: white !important;
                                    }
                                    .custom-print-footer {
                                        position: fixed;
                                        bottom: 8mm;
                                        left: 10mm;
                                        right: 10mm;
                                        display: flex;
                                        justify-content: space-between;
                                        font-size: 10px;
                                        color: #475569;
                                        font-style: italic;
                                        font-weight: 600;
                                        z-index: 1000;
                                    }
                                    @page landscapePage { size: landscape; margin: 10mm; }
                                    .print-landscape { page: landscapePage; page-break-before: always; }
                                }
                                `}
                            </style>
                            <div className="custom-print-footer hidden print:flex">
                            <span>This Document Created By {printData.type === 'expert_assignment' ? 'Tim Administrasi Teknis' : 'Tim Teknis'} Gaharu Sempana Group</span>
                        </div>

                        <table style={{ width: '100%' }}>
                            <thead>
                                <tr><td style={{ height: '12mm' }}></td></tr>
                            </thead>
                            <tbody>
                                <tr><td className="px-10 pb-4">

                                    <div className="border-b-4 border-slate-300 dark:border-slate-700 pb-4 mb-6 flex items-center justify-between">
                                        <div className="flex items-center gap-6">
                                            <img src={logoImg} alt="Gaharu Sempana Group Logo" className="h-20 object-contain" />
                                            <div>
                                                <h1 className="text-2xl font-black uppercase tracking-wider">{title}</h1>
                                                <p className="text-sm mt-1 font-semibold">SIDAMON (Sistem Database Dan Monitoring)</p>
                                                <p className="text-sm">Gaharu Sempana Group</p>
                                            </div>
                                        </div>
                                        <div className="text-right flex flex-col justify-between h-20">
                                            <div className="flex justify-end">
                                                <p className="text-xs border-2 border-slate-300 dark:border-slate-700 px-2 py-1 uppercase font-black whitespace-nowrap">Dokumen Internal</p>
                                            </div>
                                            <div className="mt-auto text-right">
                                                <p className="text-sm font-semibold whitespace-nowrap">Dicetak pada: {new Date().toLocaleDateString('id-ID', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}</p>
                                                <p className="text-xs font-semibold text-slate-700 whitespace-nowrap mt-0.5">Pukul: {new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' })}</p>
                                            </div>
                                        </div>
                                    </div>

                                    {/* BAGIAN 1: DAFTAR PROYEK */}
                                    {(printData.type === 'all' || printData.type === 'project' || (printData.type === 'custom' && printData.options.section !== 'PegawaiSaja')) && (
                                        <div className="mb-10" style={{ pageBreakAfter: (printData.type === 'all' || (printData.type === 'custom' && printData.options.section === 'Semua')) ? 'always' : 'auto' }}>
                                            <h2 className="text-lg font-bold mb-4 uppercase border-b border-gray-400 pb-2">Bagian A: Laporan Status Proyek</h2>
                                            {filteredProjects.length === 0 ? (
                                                <p className="text-sm italic">Tidak ada proyek aktif untuk kriteria ini.</p>
                                            ) : (
                                                <div className="space-y-8">
                                                    {(printData.type === 'all' || printData.type === 'custom') ? (() => {
                                                        const categoriesToRender = ((printData.type === 'custom' && printData.options.projectType !== 'Semua') ? [printData.options.projectType] : ['Perencanaan', 'Pengawasan', 'Manajemen Konstruksi']).map(kategoriTipe => {
                                                            return {
                                                                kategoriTipe,
                                                                proyeksInKategori: filteredProjects.filter(p => {
                                                                    const pType = (p.type || '').toLowerCase();
                                                                    if (kategoriTipe === 'Perencanaan') return pType.includes('perencana');
                                                                    if (kategoriTipe === 'Pengawasan') return pType === 'pengawasan' || pType === 'supervisi';
                                                                    if (kategoriTipe === 'Manajemen Konstruksi') return pType.includes('manajemen konstruksi') || pType === 'mk';
                                                                    return false;
                                                                })
                                                            };
                                                        }).filter(cat => cat.proyeksInKategori.length > 0);

                                                        return categoriesToRender.map((catData, index) => (
                                                            <div key={catData.kategoriTipe} className="mb-6" style={{ pageBreakBefore: index === 0 ? 'auto' : 'always' }}>
                                                                <h3 className="text-md font-bold mb-3 uppercase bg-gray-200 border-t border-b border-slate-300 dark:border-slate-700 py-1 px-2">Sub Bab: {catData.kategoriTipe}</h3>
                                                                <div className="space-y-6">
                                                                    {catData.proyeksInKategori.map(p => (
                                                                        <div key={p.id}>
                                                                            {renderProjectPrintCard(p)}
                                                                            {printData.type === 'project' && p.type?.toLowerCase().includes('perencana') && renderPrintTimeSchedule(p)}
                                                                        </div>
                                                                    ))}
                                                                </div>
                                                            </div>
                                                        ));
                                                    })() : (
                                                        <div className="space-y-6">
                                                            {filteredProjects.map(p => (
                                                                <div key={p.id}>
                                                                    {renderProjectPrintCard(p)}
                                                                    {printData.type === 'project' && p.type?.toLowerCase().includes('perencana') && renderPrintTimeSchedule(p)}
                                                                </div>
                                                            ))}
                                                        </div>
                                                    )}
                                                </div>
                                            )}
                                        </div>
                                    )}

                                    {/* BAGIAN 2: DAFTAR PERSONIL */}
                                    {(printData.type === 'all' || printData.type === 'personnel' || (printData.type === 'custom' && printData.options.section !== 'ProyekSaja')) && (
                                        <div>
                                            <h2 className="text-lg font-bold mb-4 uppercase border-b border-gray-400 pb-2">Bagian B: Laporan Rincian Penugasan Pegawai</h2>
                                            {filteredResources.length === 0 ? (
                                                <p className="text-sm italic">Tidak ada personil untuk kriteria ini.</p>
                                            ) : (
                                                <div className="space-y-8">
                                                    {(() => {
                                                        const resourcesBySubTeam = {};
                                                        filteredResources.forEach(res => {
                                                            const activeProjectsForRes = filteredProjects.filter(p => (p.team?.includes(res.name) || p.teamLeader === res.name || p.surveyorTeam?.includes(res.name)) && !p.notStarted);
                                                            if (printData.type === 'custom' && printData.options.projectType !== 'Semua' && activeProjectsForRes.length === 0) return;

                                                            const subTeam = getCategoryFromRole(res.role);
                                                            if (!resourcesBySubTeam[subTeam]) resourcesBySubTeam[subTeam] = [];
                                                            resourcesBySubTeam[subTeam].push({ res, activeProjectsForRes });
                                                        });

                                                        const subTeamKeys = Object.keys(resourcesBySubTeam).sort();

                                                        if (subTeamKeys.length === 0) {
                                                            return <p className="text-sm italic">Tidak ada personil yang terlibat di kriteria ini.</p>;
                                                        }

                                                        return subTeamKeys.map((subTeam, index) => (
                                                            <div key={subTeam} className="mb-6" style={{ pageBreakBefore: index === 0 ? 'auto' : 'always' }}>
                                                                <h3 className="text-md font-bold mb-3 uppercase bg-gray-200 border-t border-b border-slate-300 dark:border-slate-700 py-1 px-2">Sub Bab: Tim {subTeam}</h3>
                                                                <div className="space-y-6">
                                                                    {[...resourcesBySubTeam[subTeam]].sort((a, b) => getLPSEHierarchyScore(a.res.role) - getLPSEHierarchyScore(b.res.role) || a.res.name.localeCompare(b.res.name)).map(({ res, activeProjectsForRes }) => (
                                                                        <div key={res.id} className="border-2 border-slate-300 dark:border-slate-700 p-4" style={{ pageBreakInside: 'avoid' }}>
                                                                            <div className="flex justify-between items-center mb-2">
                                                                                <div>
                                                                                    <h3 className="font-bold uppercase text-base">{res.name}</h3>
                                                                                    {res.level && res.level.startsWith('Kordinator Divisi') ? (
                                                                                        <div className="text-xs mt-0.5 text-slate-700">
                                                                                            <p className="font-bold">{res.level}</p>
                                                                                            <p>{res.role}</p>
                                                                                        </div>
                                                                                    ) : res.level === 'Team Leader' ? (
                                                                                        <div className="text-xs mt-0.5 text-slate-700">
                                                                                            <p className="font-bold">Team Leader</p>
                                                                                            <p>{res.role}</p>
                                                                                        </div>
                                                                                    ) : res.level === 'PIC' ? (
                                                                                        <div className="text-xs mt-0.5 text-slate-700">
                                                                                            <p className="font-bold">PIC</p>
                                                                                            <p>{res.role}</p>
                                                                                        </div>
                                                                                    ) : (
                                                                                        <p className="text-xs mt-0.5 text-slate-700">
                                                                                            {res.role}
                                                                                        </p>
                                                                                    )}
                                                                                </div>
                                                                                <span className="text-xs border border-slate-300 dark:border-slate-700 px-2 py-1 font-bold">
                                                                                    {`${activeProjectsForRes.length} Proyek Aktif`}
                                                                                </span>
                                                                            </div>

                                                                            <table className="w-full text-xs border-collapse border border-slate-300 dark:border-slate-700 mt-3">
                                                                                <thead>
                                                                                    <tr className="bg-gray-100">
                                                                                        <th className="border border-slate-300 dark:border-slate-700 p-1.5 text-left w-[35%]">Nama Proyek</th>
                                                                                        <th className="border border-slate-300 dark:border-slate-700 p-1.5 text-left w-[20%]">Peran/Tim Lapangan</th>
                                                                                        <th className="border border-slate-300 dark:border-slate-700 p-1.5 text-left w-[15%]">Man Month</th>
                                                                                        <th className="border border-slate-300 dark:border-slate-700 p-1.5 text-left w-[20%]">Deadline Spesifik Tugas</th>
                                                                                        <th className="border border-slate-300 dark:border-slate-700 p-1.5 text-left w-[10%]">Status Lapangan</th>
                                                                                    </tr>
                                                                                </thead>
                                                                                <tbody>
                                                                                    {activeProjectsForRes.length === 0 ? (
                                                                                        <tr><td colSpan="5" className="border border-slate-300 dark:border-slate-700 p-1.5 text-center italic">Sedang tidak memegang proyek aktif (Available).</td></tr>
                                                                                    ) : (() => {
                                                                                        const groupedProjects = {};
                                                                                        activeProjectsForRes.forEach(p => {
                                                                                            const pType = normalizeProjectType(p.type);
                                                                                            if (!groupedProjects[pType]) groupedProjects[pType] = [];
                                                                                            groupedProjects[pType].push(p);
                                                                                        });

                                                                                        const typeKeys = Object.keys(groupedProjects).sort((a, b) => {
                                                                                            const order = { 'Perencanaan': 1, 'Pengawasan': 2, 'Manajemen Konstruksi': 3 };
                                                                                            return (order[a] || 99) - (order[b] || 99);
                                                                                        });
                                                                                        
                                                                                        return typeKeys.map(pType => (
                                                                                            <React.Fragment key={pType}>
                                                                                                <tr>
                                                                                                    <td colSpan="5" className="border border-slate-300 dark:border-slate-700 p-1.5 text-center font-bold bg-gray-200 text-[10px] uppercase">
                                                                                                        {pType}
                                                                                                    </td>
                                                                                                </tr>
                                                                                                {groupedProjects[pType].map(p => {
                                                                                                    const isPengawasan = p.type?.toLowerCase().includes('pengawas') || p.type?.toLowerCase().includes('manajemen konstruksi');
                                                                                                    let deadlineStr = '-';
                                                                                                    let roleStr = '-';
                                                                                                    let statusLapangan = '-';
                                                                                                    let manMonthStr = '-';

                                                                                                    if (p.teamLeader === res.name) {
                                                                                                        roleStr = 'Team Leader';
                                                                                                        if (isPengawasan) {
                                                                                                            deadlineStr = p.deadline ? formatDateIndo(p.deadline) : '-';
                                                                                                            const details = p.pengawasanDetails?.[res.name] || {};
                                                                                                            statusLapangan = details.statusTurun || 'Tidak Turun';
                                                                                                            manMonthStr = details.manMonth || '-';
                                                                                                        } else {
                                                                                                            statusLapangan = p.computedStatus || '-';

                                                                                                            // Kalkulasi Man Month Team Leader Perencanaan
                                                                                                            const cats = ['Arsitek', 'QS', 'Struktur', 'MEP', 'Tata Ruang', 'Surveyor', 'Lainnya'];
                                                                                                            let minDate = null;
                                                                                                            let maxDate = null;

                                                                                                            cats.forEach(cat => {
                                                                                                                const d = p.categoryDetails?.[cat];
                                                                                                                if (d) {
                                                                                                                    const startDateSource = d.startDate ? d.startDate : p.spmk;
                                                                                                                    if (startDateSource) {
                                                                                                                        const sd = new Date(startDateSource);
                                                                                                                        if (!minDate || sd < minDate) minDate = sd;
                                                                                                                    }
                                                                                                                    if (d.deadline) {
                                                                                                                        const ed = new Date(d.deadline);
                                                                                                                        if (!maxDate || ed > maxDate) maxDate = ed;
                                                                                                                    }
                                                                                                                }
                                                                                                            });

                                                                                                            if (minDate && maxDate) {
                                                                                                                minDate.setHours(0, 0, 0, 0);
                                                                                                                maxDate.setHours(0, 0, 0, 0);
                                                                                                                const diffDays = Math.ceil((maxDate.getTime() - minDate.getTime()) / (1000 * 60 * 60 * 24));
                                                                                                                if (diffDays > 0) {
                                                                                                                    manMonthStr = parseFloat((diffDays / 30).toFixed(1)).toString();
                                                                                                                }
                                                                                                            }

                                                                                                            deadlineStr = maxDate ? formatDateIndo(maxDate.toISOString().split('T')[0]) : (p.deadline ? formatDateIndo(p.deadline) : '-');
                                                                                                        }
                                                                                                    } else if (isPengawasan) {
                                                                                                        const details = p.pengawasanDetails?.[res.name] || {};
                                                                                                        deadlineStr = details.deadline ? formatDateIndo(details.deadline) : '-';
                                                                                                        roleStr = details.role || 'Inspector';
                                                                                                        statusLapangan = details.statusTurun || 'Tidak Turun';
                                                                                                        manMonthStr = details.manMonth || '-';
                                                                                                    } else {
                                                                                                        const effectiveCat = getEffectiveEmpCategory(p, res.name, res.role);
                                                                                                        const details = p.categoryDetails?.[effectiveCat] || {};
                                                                                                        deadlineStr = details.deadline ? formatDateIndo(details.deadline) : '-';
                                                                                                        roleStr = effectiveCat === 'Surveyor' ? 'Tim Surveyor' : effectiveCat;

                                                                                                        // Kalkulasi Man Month dari startDate ke deadline tim
                                                                                                        const startDateSource = details.startDate ? details.startDate : p.spmk;
                                                                                                        if (startDateSource && details.deadline) {
                                                                                                            const startD = new Date(startDateSource);
                                                                                                            const endD = new Date(details.deadline);
                                                                                                            startD.setHours(0, 0, 0, 0);
                                                                                                            endD.setHours(0, 0, 0, 0);
                                                                                                            const diffDays = Math.ceil((endD.getTime() - startD.getTime()) / (1000 * 60 * 60 * 24));
                                                                                                            if (diffDays > 0) {
                                                                                                                manMonthStr = parseFloat((diffDays / 30).toFixed(1)).toString();
                                                                                                            }
                                                                                                        }

                                                                                                        // Status Lapangan diambil dari status sub tim
                                                                                                        statusLapangan = getMicroStatus(details.progress || 0, details.deadline);
                                                                                                    }

                                                                                                    return (
                                                                                                        <tr key={p.id}>
                                                                                                            <td className="border border-slate-300 dark:border-slate-700 p-1.5">{p.name}</td>
                                                                                                            <td className="border border-slate-300 dark:border-slate-700 p-1.5">{roleStr}</td>
                                                                                                            <td className="border border-slate-300 dark:border-slate-700 p-1.5 text-center">{manMonthStr}</td>
                                                                                                            <td className="border border-slate-300 dark:border-slate-700 p-1.5 font-semibold">{deadlineStr}</td>
                                                                                                            <td className="border border-slate-300 dark:border-slate-700 p-1.5 font-semibold">{statusLapangan}</td>
                                                                                                        </tr>
                                                                                                    )
                                                                                                })}
                                                                                            </React.Fragment>
                                                                                        ));
                                                                                    })()}
                                                                                </tbody>
                                                                            </table>
                                                                        </div>
                                                                    ))}
                                                                </div>
                                                            </div>
                                                        ));
                                                    })()}
                                                </div>
                                            )}
                                        </div>
                                    )}

                                    {/* BAGIAN 3: DAFTAR PENUGASAN TENAGA AHLI */}
                                    {(printData.type === 'expert_assignment') && (() => {
                                        const activeAsg = assignments.filter(asg => {
                                            if (asg.endDate) {
                                                const today = new Date();
                                                today.setHours(0, 0, 0, 0);
                                                const end = new Date(asg.endDate);
                                                end.setHours(0, 0, 0, 0);
                                                if (end < today) return false;
                                            }
                                            return true;
                                        });

                                        return (
                                            <div>
                                                <h2 className="text-lg font-bold mb-4 uppercase border-b border-gray-400 pb-2">Laporan Rekam Penugasan Tenaga Ahli</h2>
                                                {activeAsg.length === 0 ? (
                                                    <p className="text-sm italic">Tidak ada penugasan tenaga ahli aktif yang terdaftar.</p>
                                                ) : (
                                                    <div className="space-y-6">
                                                        {activeAsg.map(asg => (
                                                            <div key={asg.id} className="border-2 border-slate-300 dark:border-slate-700 p-4" style={{ pageBreakInside: 'avoid' }}>
                                                                <div className="flex justify-between items-center mb-2">
                                                                    <div>
                                                                        <h3 className="font-bold uppercase text-base">{asg.jobName}</h3>
                                                                        <p className="text-xs uppercase mt-0.5 font-bold">{asg.tenderType} | {asg.lpseName}</p>
                                                                        <p className="text-xs mt-0.5">Tipe Proyek & Kontrak: {asg.projectType || 'Pengawasan'} - {asg.contractType}</p>
                                                                        <p className="text-xs mt-0.5">SPMK - Berakhir: {asg.startDate ? formatDateIndo(asg.startDate) : '-'} s/d {asg.endDate ? formatDateIndo(asg.endDate) : '-'} ({asg.duration || 0} Hari)</p>
                                                                    </div>
                                                                </div>

                                                                <table className="w-full text-xs border-collapse border border-slate-300 dark:border-slate-700 mt-3">
                                                                    <thead>
                                                                        <tr className="bg-gray-100">
                                                                            <th className="border border-slate-300 dark:border-slate-700 p-1.5 text-left w-[40%]">Nama</th>
                                                                            <th className="border border-slate-300 dark:border-slate-700 p-1.5 text-left w-[30%]">Sertifikat (SKA)</th>
                                                                            <th className="border border-slate-300 dark:border-slate-700 p-1.5 text-center w-[15%]">Peran</th>
                                                                            <th className="border border-slate-300 dark:border-slate-700 p-1.5 text-center w-[15%]">Man Month</th>
                                                                        </tr>
                                                                    </thead>
                                                                    <tbody>
                                                                        {(asg.experts || []).length === 0 ? (
                                                                            <tr><td colSpan="4" className="border border-slate-300 dark:border-slate-700 p-1.5 text-center italic">Belum ada tenaga ahli.</td></tr>
                                                                        ) : [...(asg.experts || [])].sort((a, b) => getLPSEHierarchyScore(a.role) - getLPSEHierarchyScore(b.role) || (experts.find(e => e.id === a.expertId)?.name || '').localeCompare(experts.find(e => e.id === b.expertId)?.name || '')).map((expPlot, idx) => {
                                                                            const exp = experts.find(e => e.id === expPlot.expertId);
                                                                            const expertName = exp ? exp.name : 'Unknown';
                                                                            const certObj = exp ? (exp.certificates || []).find(c => c.certName === expPlot.certificateName) : null;
                                                                            let certDisplay = certObj && certObj.certLevel ? `${expPlot.certificateName} (${certObj.certLevel})` : (expPlot.certificateName || '-');
                                                                            
                                                                            (expPlot.additionalCertificates || []).forEach(addCert => {
                                                                                if (!addCert) return;
                                                                                const cObj = exp ? (exp.certificates || []).find(c => c.certName === addCert) : null;
                                                                                const cDisplay = cObj && cObj.certLevel ? `${addCert} (${cObj.certLevel})` : addCert;
                                                                                if (certDisplay === '-') certDisplay = cDisplay;
                                                                                else certDisplay += ` & ${cDisplay}`;
                                                                            });
                                                                            
                                                                            return (
                                                                                <tr key={idx}>
                                                                                    <td className="border border-slate-300 dark:border-slate-700 p-1.5 font-semibold">{expertName}</td>
                                                                                    <td className="border border-slate-300 dark:border-slate-700 p-1.5">{certDisplay}</td>
                                                                                    <td className="border border-slate-300 dark:border-slate-700 p-1.5 text-center">{expPlot.role || '-'}</td>
                                                                                    <td className="border border-slate-300 dark:border-slate-700 p-1.5 text-center font-bold">{expPlot.manMonth}</td>
                                                                                </tr>
                                                                            )
                                                                        })}
                                                                    </tbody>
                                                                </table>
                                                            </div>
                                                        ))}
                                                    </div>
                                                )}
                                            </div>
                                        );
                                    })()}

                                    {/* BAGIAN TANDA TANGAN */}
                                    <div className="mt-16 pt-8 flex justify-between px-10 text-sm text-slate-800 dark:text-slate-200" style={{ pageBreakInside: 'avoid' }}>
                                        <div className="text-center">
                                            <p className="mb-24">Diketahui Oleh,<br />{printData.type === 'expert_assignment' ? 'Manajer Administrasi' : 'Manajer Teknis'} Gaharu Sempana Group</p>
                                            <p className="font-bold underline decoration-black underline-offset-2">{printData.type === 'expert_assignment' ? 'Gusti Ayu Legong Aryaningsih, S.T., M.T.' : 'Ar. Ir. I Nyoman Adi Putra Wijaya, ST. IAI.'}</p>
                                        </div>
                                        <div className="text-center">
                                            <p className="mb-24">Disetujui Oleh,<br />CEO Gaharu Sempana Group</p>
                                            <p className="font-bold underline decoration-black underline-offset-2">Ir. Putu Andre Wicaksana Putra, S.T., M.Ars., IPP.</p>
                                        </div>
                                    </div>

                                    {/* --- AKHIR KONTEN UTAMA --- */}
                                </td></tr>
                            </tbody>
                            <tfoot>
                                <tr><td style={{ height: '20mm' }}></td></tr>
                            </tfoot>
                        </table>
                        </div>
                    </div>
                );
            };
            // --- KOMPONEN TAB: TIME SCHEDULE (GANTT CHART) ---
            const renderTimeSchedule = () => {
                if (!activeScheduleProject) return null;
                const p = activeScheduleProject;

                // 1. Ekstrak Fase/Tim yang aktif di proyek ini
                const phases = [];
                const possibleCats = ['Arsitek', 'Struktur', 'MEP', 'Tata Ruang', 'QS', 'Surveyor', 'Lainnya'];

                let globalMinDate = null;
                let globalMaxDate = null;

                const isPerencanaan = p.type?.toLowerCase().includes('perencana');

                possibleCats.forEach(cat => {
                    const detail = p.categoryDetails?.[cat];
                    const hasData = detail && (detail.startDate || detail.deadline || (detail.progress !== undefined && detail.progress !== ""));
                    const hasSpmkData = isPerencanaan && p.spmk && detail && (detail.deadline || (detail.progress !== undefined && detail.progress !== ""));

                    // Tampilkan jika punya startDate/deadline atau progress > 0
                    if (hasData || hasSpmkData) {
                        const rawStartDate = detail.startDate ? detail.startDate : p.spmk;
                        let sd = rawStartDate ? new Date(rawStartDate) : null;
                        let ed = detail.deadline ? new Date(detail.deadline) : null;

                        if (sd && (!globalMinDate || sd < globalMinDate)) globalMinDate = new Date(sd);
                        if (ed && (!globalMaxDate || ed > globalMaxDate)) globalMaxDate = new Date(ed);
                        if (sd && !globalMaxDate) globalMaxDate = new Date(sd);
                        if (ed && !globalMinDate) globalMinDate = new Date(ed);

                        // Cari assignees (orang yang perannya masuk ke kategori ini dan ada di p.team)
                        const assignees = p.team ? p.team.filter(mName => {
                            const res = resources.find(r => r.name === mName);
                            if (res) {
                                return getEffectiveEmpCategory(p, mName, res.role) === cat;
                            }
                            return false;
                        }) : [];

                        if (cat === 'Surveyor' && p.surveyorTeam && p.surveyorTeam.length > 0) {
                            p.surveyorTeam.forEach(st => { if (!assignees.includes(st)) assignees.push(st); });
                        }

                        phases.push({
                            id: cat,
                            name: cat,
                            startDate: sd,
                            endDate: ed,
                            progress: detail.progress || 0,
                            assignees: assignees,
                            isSubTask: false
                        });

                        if (detail.tasks && detail.tasks.length > 0) {
                            detail.tasks.forEach((task, tIdx) => {
                                const isTaskCompleted = (detail.completedTasks || []).includes(task);
                                phases.push({
                                    id: `${cat}-task-${tIdx}`,
                                    name: task,
                                    startDate: null,
                                    endDate: null,
                                    progress: isTaskCompleted ? 100 : 0,
                                    assignees: [],
                                    isSubTask: true
                                });
                            });
                        }
                    }
                });

                // Jika tidak ada data tanggal sama sekali, gunakan hari ini sebagai default
                if (!globalMinDate) globalMinDate = new Date();
                if (!globalMaxDate) {
                    globalMaxDate = new Date();
                    globalMaxDate.setMonth(globalMaxDate.getMonth() + 3); // Default span 3 bulan
                }

                // Beri padding 1 bulan ke belakang dan 2 bulan ke depan untuk visual
                const chartStart = new Date(globalMinDate);
                chartStart.setDate(1);

                const chartEnd = new Date(globalMaxDate);
                chartEnd.setMonth(chartEnd.getMonth() + 2);
                chartEnd.setDate(0); // Akhir bulan

                // Generate array of months for the header
                const timeUnits = [];
                let curr = new Date(chartStart);

                if (scheduleZoom === 'month') {
                    while (curr <= chartEnd) {
                        timeUnits.push({ date: new Date(curr), label: curr.toLocaleDateString('id-ID', { month: 'short', year: '2-digit' }).toUpperCase(), monthLabel: curr.toLocaleDateString('id-ID', { month: 'short', year: '2-digit' }).toUpperCase() });
                        curr.setMonth(curr.getMonth() + 1);
                    }
                } else {
                    let currentMonth = -1;
                    let weekCounter = 1;
                    while (curr <= chartEnd) {
                        const monthStr = curr.toLocaleDateString('id-ID', { month: 'long', year: 'numeric' }).toUpperCase();
                        if (curr.getMonth() !== currentMonth) {
                            currentMonth = curr.getMonth();
                            weekCounter = 1;
                        }
                        timeUnits.push({ date: new Date(curr), label: `MG ${weekCounter}`, monthLabel: monthStr });

                        let nextDate = new Date(curr);
                        nextDate.setDate(nextDate.getDate() + 7);

                        if (nextDate.getMonth() !== currentMonth) {
                            curr = new Date(curr.getFullYear(), curr.getMonth() + 1, 1);
                        } else {
                            curr = nextDate;
                        }
                        weekCounter++;
                    }
                }
                const totalDays = Math.round((chartEnd - chartStart) / (1000 * 60 * 60 * 24)) + 1;

                timeUnits.forEach((u, i) => {
                    let nextDate = i < timeUnits.length - 1 ? timeUnits[i + 1].date : new Date(chartEnd.getTime() + 1000 * 60 * 60 * 24);
                    u.durationDays = Math.round((nextDate - u.date) / (1000 * 60 * 60 * 24));
                    u.widthPercent = (u.durationDays / totalDays) * 100;
                });



                // Filter phases by user
                let filteredPhases = phases;
                if (scheduleFilterUser !== 'Semua') {
                    filteredPhases = phases.filter(ph => {
                        if (ph.isSubTask) {
                            const parent = phases.find(parentPhase => parentPhase.id === ph.id.split('-task')[0]);
                            return parent && parent.assignees && parent.assignees.includes(scheduleFilterUser);
                        }
                        return ph.assignees && ph.assignees.includes(scheduleFilterUser);
                    });
                }

                // Filter out collapsed subtasks
                filteredPhases = filteredPhases.filter(ph => {
                    if (ph.isSubTask) {
                        const parentId = ph.id.split('-task')[0];
                        if (scheduleCollapsedCats[parentId]) return false;
                    }
                    return true;
                });

                // Warna ditentukan secara dinamis berdasarkan status fase (Selesai, Beresiko, Terlambat, On Progress)

                return (
                    <div className="flex flex-col h-[calc(100vh-12rem)] min-h-[500px] fade-in glass-card rounded-[2rem] overflow-hidden">
                        {/* Header Area */}
                        <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 px-6 py-4 border-b border-slate-100 dark:border-slate-700/50 bg-white/40 dark:bg-slate-800/40 shrink-0">
                            <div className="flex items-center gap-3">
                                <button onClick={() => { setActiveTab('proyek'); setActiveScheduleProject(null); }} className="p-2 rounded-full hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors text-slate-500">
                                    <Icon name="arrow-left" size={20} />
                                </button>
                                <div className="flex items-center gap-2 text-sm">
                                    <span className="text-slate-500 font-medium cursor-pointer hover:text-slate-800 dark:hover:text-slate-300" onClick={() => { setActiveTab('proyek'); setActiveScheduleProject(null); }}>Projects</span>
                                    <span className="text-slate-400">/</span>
                                    <span className="text-slate-500 font-medium">Perencanaan</span>
                                    <span className="text-slate-400">/</span>
                                    <span className="text-slate-800 dark:text-slate-200 font-bold">{p.name}</span>
                                </div>
                            </div>
                            <div className="flex items-center gap-2">
                                <select
                                    className="px-3 py-1.5 text-xs font-semibold rounded-lg bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 text-slate-700 dark:text-slate-300 outline-none focus:ring-2 focus:ring-purple-500/50"
                                    value={scheduleZoom}
                                    onChange={(e) => setScheduleZoom(e.target.value)}
                                >
                                    <option value="month">Tampilan Bulanan</option>
                                    <option value="week">Tampilan Mingguan</option>
                                </select>
                                <select
                                    className="px-3 py-1.5 text-xs font-semibold rounded-lg bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 text-slate-700 dark:text-slate-300 outline-none focus:ring-2 focus:ring-purple-500/50"
                                    value={scheduleFilterUser}
                                    onChange={(e) => setScheduleFilterUser(e.target.value)}
                                >
                                    <option value="Semua">Semua Personil</option>
                                    {(p.team || []).map(m => (
                                        <option key={m} value={m}>{m}</option>
                                    ))}
                                </select>

                            </div>
                        </div>

                        {/* Sub Header (Timeline Span & Members) */}
                        <div className="px-6 py-3 border-b border-slate-100 dark:border-slate-800/50 flex justify-between items-center bg-slate-50/50 dark:bg-slate-900/30 shrink-0">
                            <div className="flex items-center gap-3 text-xs">
                                <span className="font-semibold text-slate-500">Team Leader: <span className="text-slate-800 dark:text-slate-200 font-bold">{p.teamLeader || "Belum ditugaskan"}</span></span>
                                <span className="w-1 h-1 rounded-full bg-slate-300"></span>
                                <span className="font-semibold text-slate-500">Timeline: </span>
                                <span className="bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 px-3 py-1 rounded-full text-slate-700 dark:text-slate-300 font-medium shadow-sm flex items-center gap-2">
                                    {formatDateIndo(globalMinDate.toISOString().split('T')[0])} to {formatDateIndo(globalMaxDate.toISOString().split('T')[0])}
                                    <Icon name="chevrons-up-down" size={12} className="text-slate-400" />
                                </span>
                            </div>
                        </div>

                        {/* Gantt Area */}
                        <div className="flex-1 overflow-auto flex bg-white/40 dark:bg-slate-900/40">

                            {/* Left Column: Items Table */}
                            <div className="w-72 sm:w-80 border-r border-slate-200 dark:border-slate-800 shrink-0 flex flex-col bg-white dark:bg-slate-950 z-10 sticky left-0 shadow-[4px_0_12px_-4px_rgba(0,0,0,0.05)] dark:shadow-none">
                                <div className="h-10 border-b border-slate-200 dark:border-slate-800 flex items-center px-4 shrink-0 bg-slate-50 dark:bg-slate-900/50">
                                    <span className="text-xs font-bold text-slate-500 uppercase tracking-wider">Tasks</span>
                                </div>
                                <div className="flex-1 overflow-y-auto">
                                    {filteredPhases.length === 0 ? (
                                        <div className="p-4 text-xs text-slate-400 italic">Tidak ada data fase / jadwal untuk kriteria ini.</div>
                                    ) : (
                                        <div className="flex flex-col">
                                            {filteredPhases.map((phase) => {
                                                const isDone = phase.progress >= 100;
                                                return (
                                                    <div key={phase.id} className={`h-12 border-b border-slate-100 dark:border-slate-800/50 flex items-center px-4 gap-2 hover:bg-slate-50 dark:hover:bg-slate-900/30 transition-colors group ${phase.isSubTask ? 'pl-10 bg-slate-50/30 dark:bg-slate-900/10' : ''}`}>
                                                        {!phase.isSubTask ? (
                                                            <button
                                                                onClick={() => setScheduleCollapsedCats(prev => ({ ...prev, [phase.id]: !prev[phase.id] }))}
                                                                className="text-slate-400 hover:text-indigo-500 transition-colors"
                                                                title={scheduleCollapsedCats[phase.id] ? "Buka Sub-Tugas" : "Tutup Sub-Tugas"}
                                                            >
                                                                <Icon name={scheduleCollapsedCats[phase.id] ? "chevron-right" : "chevron-down"} size={16} />
                                                            </button>
                                                        ) : (
                                                            <div className="w-3 border-b-2 border-l-2 border-slate-300 dark:border-slate-600 h-3 rounded-bl -mt-2 ml-1 opacity-50 shrink-0"></div>
                                                        )}
                                                        <div className={`w-4 h-4 rounded shadow-sm border flex items-center justify-center shrink-0 ${isDone ? 'bg-emerald-500 border-emerald-600' : 'bg-white dark:bg-slate-800 border-slate-300 dark:border-slate-600'}`}>
                                                            {isDone && <Icon name="check" size={10} className="text-white" strokeWidth={3} />}
                                                        </div>
                                                        <span className={`${phase.isSubTask ? 'text-xs' : 'text-sm'} font-semibold truncate ${isDone ? 'text-slate-400 dark:text-slate-500' : 'text-slate-700 dark:text-slate-300'}`} title={phase.name}>{phase.name}</span>
                                                        {isDone && <Icon name="check-circle-2" size={14} className="text-emerald-500 ml-auto shrink-0" />}
                                                    </div>
                                                )
                                            })}
                                        </div>
                                    )}
                                </div>
                            </div>

                            {/* Right Column: Timeline Grid */}
                            <div className="flex-1 flex flex-col relative min-w-[600px] overflow-x-auto">
                                {/* Months Header */}
                                <div className="border-b border-slate-200 dark:border-slate-800 flex flex-col relative shrink-0 bg-slate-50 dark:bg-slate-900/50 sticky top-0 z-10 w-full" style={{ minWidth: '100%' }}>
                                    {scheduleZoom === 'week' && (
                                        <div className="flex w-full h-5 border-b border-slate-200 dark:border-slate-800 bg-slate-200/50 dark:bg-slate-800/80">
                                            {/* Group by monthLabel */}
                                            {(() => {
                                                const monthGroups = [];
                                                timeUnits.forEach(u => {
                                                    const last = monthGroups[monthGroups.length - 1];
                                                    if (last && last.monthLabel === u.monthLabel) {
                                                        last.widthPercent += u.widthPercent;
                                                    } else {
                                                        monthGroups.push({ monthLabel: u.monthLabel, widthPercent: u.widthPercent });
                                                    }
                                                });
                                                return monthGroups.map((g, i) => (
                                                    <div key={`m-${i}`} className="flex items-center justify-center border-l border-slate-300 dark:border-slate-700 first:border-l-0 text-slate-700 dark:text-slate-300 font-bold tracking-wider text-[10px] overflow-hidden whitespace-nowrap text-ellipsis" style={{ width: `${g.widthPercent}%` }}>
                                                        {g.monthLabel}
                                                    </div>
                                                ));
                                            })()}
                                        </div>
                                    )}
                                    <div className={`flex w-full ${scheduleZoom === 'week' ? 'h-5' : 'h-10'}`}>
                                        {timeUnits.map((u, i) => (
                                            <div key={`u-${i}`} className={`flex items-center justify-center border-l border-slate-200 dark:border-slate-800/80 first:border-l-0 text-[10px] font-semibold ${scheduleZoom === 'month' ? 'text-slate-600 dark:text-slate-400 tracking-wider' : 'text-slate-400 dark:text-slate-500'} overflow-hidden whitespace-nowrap text-ellipsis`} style={{ width: `${u.widthPercent}%` }}>
                                                {u.label}
                                            </div>
                                        ))}
                                    </div>
                                </div>

                                {/* Grid Background Lines */}
                                <div className="absolute top-10 bottom-0 left-0 right-0 flex pointer-events-none z-0">
                                    {timeUnits.map((m, i) => (
                                        <div key={`grid-${i}`} className="border-l border-slate-100 dark:border-slate-800/30 first:border-l-0 border-dashed" style={{ width: `${m.widthPercent}%` }}></div>
                                    ))}
                                </div>

                                {/* Bars Area */}
                                <div className="flex-1 w-full relative z-10">
                                    {filteredPhases.map((phase, idx) => {
                                        let hasBar = false;
                                        let leftPercent = 0;
                                        let widthPercent = 0;
                                        let isLate = false;
                                        let phaseStatus = "On Progress";

                                        if (phase.startDate && phase.endDate) {
                                            hasBar = true;
                                            const startOffset = Math.max(0, (phase.startDate - chartStart) / (1000 * 60 * 60 * 24));
                                            const duration = (phase.endDate - phase.startDate) / (1000 * 60 * 60 * 24);
                                            leftPercent = (startOffset / totalDays) * 100;
                                            widthPercent = Math.max(1, (duration / totalDays) * 100);

                                            // Get precise micro status
                                            phaseStatus = getMicroStatus(phase.progress, phase.endDate);
                                            if (phaseStatus === "Terlambat") {
                                                isLate = true;
                                            }
                                        } else if (phase.progress >= 100) {
                                            phaseStatus = "Done";
                                        }

                                        let colorClass = "bg-blue-500 border-blue-600 text-white"; // On Progress (Biru)
                                        if (phaseStatus === "Done") {
                                            colorClass = "bg-emerald-500 border-emerald-600 text-white"; // Hijau
                                        } else if (phaseStatus === "Terlambat") {
                                            colorClass = "bg-red-500 border-red-600 text-white"; // Merah
                                        } else if (phaseStatus === "Beresiko") {
                                            colorClass = "bg-yellow-400 border-yellow-500 text-slate-800"; // Kuning
                                        }



                                        return (
                                            <div key={`bar-${phase.id}`} className="h-12 border-b border-slate-100/50 dark:border-slate-800/30 flex items-center relative hover:bg-slate-50/50 dark:hover:bg-slate-900/10 transition-colors w-full">
                                                {hasBar && (
                                                    <div
                                                        className={`absolute h-6 rounded-md border flex items-center shadow-sm group cursor-pointer hover:shadow-md transition-shadow hover:z-50 ${colorClass}`}
                                                        style={{ left: `${leftPercent}%`, width: chartAnimate ? `${widthPercent}%` : '0%', minWidth: chartAnimate ? '40px' : '0px', transition: 'width 1s cubic-bezier(0.16, 1, 0.3, 1)' }}
                                                    >
                                                        {/* Progress fill container */}
                                                        <div className="absolute inset-0 overflow-hidden rounded-md pointer-events-none">
                                                            <div className={`absolute top-0 bottom-0 left-0 opacity-20 ${isLate ? 'bg-black' : 'bg-current'}`} style={{ width: `${phase.progress}%` }}></div>
                                                        </div>

                                                        {/* Label inside bar */}
                                                        <div className="px-2 flex items-center gap-1.5 z-10 w-full h-full text-[10px] font-bold pointer-events-none overflow-hidden">
                                                            {isLate && <Icon name="alert-triangle" size={10} className="text-white shrink-0" />}
                                                            {phase.progress === 100 && <Icon name="check-circle-2" size={10} className="shrink-0" />}
                                                        </div>

                                                        {/* Rich Tooltip (Glassmorphism) */}
                                                        <div className="absolute top-full left-0 mt-2 hidden group-hover:block z-[100] w-64 bg-slate-900/90 backdrop-blur-md text-white rounded-xl p-3 shadow-2xl pointer-events-none border border-slate-700">
                                                            <div className="font-bold text-sm mb-1">{phase.name}</div>
                                                            <div className="flex justify-between items-center text-xs mb-1">
                                                                <span className="text-slate-400">Progress</span>
                                                                <span className="font-bold">{phase.progress}%</span>
                                                            </div>
                                                            <div className="h-1.5 w-full bg-slate-800 rounded-full mb-3 overflow-hidden shadow-inner">
                                                                <div className="h-full bg-emerald-500 rounded-full" style={{ width: `${phase.progress}%` }}></div>
                                                            </div>
                                                            <div className="flex flex-col gap-1.5">
                                                                <div className="flex items-center gap-2 text-[10px] text-slate-300">
                                                                    <Icon name="calendar" size={12} className="text-slate-400 shrink-0" />
                                                                    {formatDateIndo(new Date(phase.startDate).toISOString().split('T')[0])} - {formatDateIndo(new Date(phase.endDate).toISOString().split('T')[0])}
                                                                </div>
                                                                <div className="flex items-center gap-2 text-[10px] text-slate-300">
                                                                    <Icon name="clock" size={12} className="text-slate-400 shrink-0" />
                                                                    {(() => {
                                                                        if (phaseStatus === 'Done') return 'Selesai';
                                                                        const diffDays = Math.ceil((new Date(phase.endDate) - new Date()) / (1000 * 60 * 60 * 24));
                                                                        if (diffDays < 0) return <span className="text-red-400 font-bold">Terlambat {Math.abs(diffDays)} hari</span>;
                                                                        return `Tersisa ${diffDays} hari`;
                                                                    })()}
                                                                </div>
                                                                <div className="flex gap-2 text-[10px] text-slate-300 mt-1 pt-1.5 border-t border-slate-700/50">
                                                                    <Icon name="users" size={12} className="text-slate-400 shrink-0 mt-0.5" />
                                                                    <span className="leading-tight">{phase.assignees && phase.assignees.length > 0 ? phase.assignees.join(', ') : 'Belum ditugaskan'}</span>
                                                                </div>

                                                            </div>
                                                        </div>


                                                    </div>
                                                )}
                                            </div>
                                        )
                                    })}

                                    {/* Today Line */}
                                    {(() => {
                                        const today = new Date();
                                        if (today >= chartStart && today <= chartEnd) {
                                            const todayOffset = (today - chartStart) / (1000 * 60 * 60 * 24);
                                            const todayPercent = (todayOffset / totalDays) * 100;
                                            return (
                                                <div
                                                    className="absolute top-0 bottom-0 border-l-2 border-rose-500 border-dashed z-20 pointer-events-none"
                                                    style={{ left: `${todayPercent}%` }}
                                                >
                                                    <div className="bg-rose-500 text-white text-[9px] font-bold px-2 py-0.5 rounded-full absolute -top-3 -translate-x-1/2 shadow-[0_0_15px_rgba(244,63,94,0.6)] animate-pulse whitespace-nowrap flex items-center gap-1">
                                                        <div className="w-1.5 h-1.5 bg-white rounded-full"></div> Today
                                                    </div>
                                                </div>
                                            );
                                        }
                                        return null;
                                    })()}

                                </div>
                            </div>
                        </div>
                    </div>
                );
            };

            const renderMasterSchedule = () => {
                let displayedProjects = projects.filter(p => p.spmk && p.deadline && p.type?.toLowerCase().includes('perencanaan') && Number(p.progress || 0) < 100 && p.status !== 'Done');

                if (displayedProjects.length === 0) {
                    return (
                        <div className="relative flex flex-col h-[calc(100vh-12rem)] min-h-[400px] bg-white/80 dark:bg-slate-800/80 backdrop-blur-3xl rounded-[2rem] border border-slate-100 dark:border-slate-700/50 shadow-[0_8px_30px_rgb(0,0,0,0.04)] items-center justify-center p-8 text-center">
                            <div className="absolute inset-0 bg-[radial-gradient(ellipse_at_center,_var(--tw-gradient-stops))] from-indigo-50/50 via-transparent to-transparent"></div>
                            <div className="relative z-10 flex flex-col items-center">
                                <div className="w-20 h-20 bg-indigo-100 dark:bg-indigo-900/30 rounded-full flex items-center justify-center mb-6">
                                    <Icon name="folder-open" size={40} className="text-indigo-500" />
                                </div>
                                <h3 className="text-xl font-bold text-slate-800 dark:text-white mb-2">Master Schedule Kosong</h3>
                                <p className="text-sm text-slate-500 dark:text-slate-400 max-w-sm">Tidak ada data proyek dengan tanggal SPMK dan Deadline yang aktif saat ini.</p>
                            </div>
                        </div>
                    );
                }

                let globalMinDate = new Date(Math.min(...displayedProjects.map(p => new Date(p.spmk))));
                let globalMaxDate = new Date(Math.max(...displayedProjects.map(p => new Date(p.deadline))));

                if (isNaN(globalMinDate.getTime())) globalMinDate = new Date();
                if (isNaN(globalMaxDate.getTime())) globalMaxDate = new Date(globalMinDate.getTime() + 30 * 24 * 60 * 60 * 1000);

                let chartStart = new Date(globalMinDate);
                chartStart.setMonth(chartStart.getMonth() - 1);
                chartStart.setDate(1);

                let chartEnd = new Date(globalMaxDate);
                chartEnd.setMonth(chartEnd.getMonth() + 2);
                chartEnd.setDate(0);

                const timeUnits = [];
                let curr = new Date(chartStart);
                if (scheduleZoom === 'month') {
                    while (curr <= chartEnd) {
                        timeUnits.push({ date: new Date(curr), label: curr.toLocaleDateString('id-ID', { month: 'short', year: '2-digit' }).toUpperCase(), monthLabel: curr.toLocaleDateString('id-ID', { month: 'long', year: 'numeric' }).toUpperCase() });
                        curr.setMonth(curr.getMonth() + 1);
                    }
                } else {
                    let currentMonth = -1;
                    let weekCounter = 1;
                    while (curr <= chartEnd) {
                        const monthStr = curr.toLocaleDateString('id-ID', { month: 'long', year: 'numeric' }).toUpperCase();
                        if (curr.getMonth() !== currentMonth) {
                            currentMonth = curr.getMonth();
                            weekCounter = 1;
                        }
                        timeUnits.push({ date: new Date(curr), label: `MG ${weekCounter}`, monthLabel: monthStr });

                        let nextDate = new Date(curr);
                        nextDate.setDate(nextDate.getDate() + 7);

                        if (nextDate.getMonth() !== currentMonth) {
                            curr = new Date(curr.getFullYear(), curr.getMonth() + 1, 1);
                        } else {
                            curr = nextDate;
                        }
                        weekCounter++;
                    }
                }

                const totalDays = Math.round((chartEnd - chartStart) / (1000 * 60 * 60 * 24)) + 1;

                timeUnits.forEach((u, i) => {
                    let nextDate = i < timeUnits.length - 1 ? timeUnits[i + 1].date : new Date(chartEnd.getTime() + 1000 * 60 * 60 * 24);
                    u.durationDays = Math.round((nextDate - u.date) / (1000 * 60 * 60 * 24));
                    u.widthPercent = 100 / timeUnits.length;
                });

                const getDatePercent = (dateObj) => {
                    if (dateObj < chartStart) return 0;
                    if (dateObj > chartEnd) return 100;
                    for (let i = 0; i < timeUnits.length; i++) {
                        const u = timeUnits[i];
                        const nextDate = i < timeUnits.length - 1 ? timeUnits[i + 1].date : new Date(chartEnd.getTime() + 1000 * 60 * 60 * 24);
                        if (dateObj >= u.date && dateObj <= nextDate) {
                            const fraction = (dateObj - u.date) / (nextDate - u.date);
                            return (i * (100 / timeUnits.length)) + (fraction * (100 / timeUnits.length));
                        }
                    }
                    return 100;
                };

                return (
                    <div className="flex flex-col h-[calc(100vh-12rem)] min-h-[500px] fade-in glass-card rounded-[2rem] overflow-hidden">
                        <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 px-6 py-4 border-b border-slate-100 dark:border-slate-700/50 bg-white/40 dark:bg-slate-800/40 shrink-0">
                            <div className="flex items-center gap-3">
                                <div className="flex items-center gap-2 text-sm">
                                    <span className="text-slate-800 dark:text-slate-100 font-bold text-lg">Master Schedule (Portofolio Proyek)</span>
                                </div>
                            </div>
                            <div className="flex items-center gap-2">
                                <select
                                    className="px-3 py-1.5 text-xs font-semibold rounded-lg bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 text-slate-700 dark:text-slate-300 outline-none focus:ring-2 focus:ring-purple-500/50"
                                    value={scheduleZoom}
                                    onChange={(e) => setScheduleZoom(e.target.value)}
                                >
                                    <option value="month">Tampilan Bulanan</option>
                                    <option value="week">Tampilan Mingguan</option>
                                </select>
                            </div>
                        </div>

                        <div className="flex-1 overflow-auto relative bg-white dark:bg-slate-950 fade-in">
                            <div className="flex min-w-max min-h-full">
                                {/* Left Column: Project Names */}
                                <div className="w-80 border-r border-slate-200 dark:border-slate-800 shrink-0 flex flex-col bg-white dark:bg-slate-950 z-30 sticky left-0 shadow-[4px_0_12px_-4px_rgba(0,0,0,0.05)] dark:shadow-none">
                                    <div className="h-10 border-b border-slate-200 dark:border-slate-800 flex items-center px-4 shrink-0 bg-slate-50 dark:bg-slate-900/50 sticky top-0 z-40">
                                        <span className="text-xs font-bold text-slate-500 uppercase tracking-wider">Daftar Proyek</span>
                                    </div>
                                    <div className="flex flex-col">
                                        {displayedProjects.map((p) => {
                                            const isDone = p.status === 'Done';
                                            return (
                                                <div key={p.id} className="h-14 border-b border-slate-100 dark:border-slate-800/50 flex items-center px-4 hover:bg-slate-50 dark:hover:bg-slate-900/30 transition-colors">
                                                    <div className="flex flex-col min-w-0">
                                                        <span className="text-sm font-bold text-slate-700 dark:text-slate-300 line-clamp-2 leading-tight" title={p.name}>{p.name}</span>
                                                        <span className="text-[10px] text-slate-500 font-semibold mt-0.5">
                                                            {p.type?.toLowerCase().includes('perencana') ? 'Perencanaan' : p.type} • Akhir Kontrak: {formatDateIndo(p.deadline)}
                                                        </span>
                                                    </div>
                                                </div>
                                            )
                                        })}
                                    </div>
                                </div>

                                {/* Right Column: Gantt Chart */}
                                <div className="flex-1 flex flex-col relative bg-slate-50/30 dark:bg-slate-950" style={{ minWidth: scheduleZoom === 'month' ? `${timeUnits.length * 80}px` : `${timeUnits.length * 40}px` }}>
                                    <div className="flex flex-col border-b border-slate-200 dark:border-slate-800 shrink-0 sticky top-0 z-20 bg-slate-50 dark:bg-slate-900/80 backdrop-blur-sm">
                                        {scheduleZoom === 'week' && (
                                            <div className="flex h-5 border-b border-slate-200 dark:border-slate-700">
                                                {(() => {
                                                    const monthGroups = [];
                                                    timeUnits.forEach(u => {
                                                        const last = monthGroups[monthGroups.length - 1];
                                                        if (last && last.monthLabel === u.monthLabel) last.widthPercent += u.widthPercent;
                                                        else monthGroups.push({ monthLabel: u.monthLabel, widthPercent: u.widthPercent });
                                                    });
                                                    return monthGroups.map((g, i) => (
                                                        <div key={`m-${i}`} className="flex items-center justify-center border-r border-slate-200 dark:border-slate-700 text-[10px] font-bold text-slate-500 tracking-widest uppercase bg-slate-100/50 dark:bg-slate-800/50 overflow-hidden whitespace-nowrap text-ellipsis" style={{ width: `${g.widthPercent}%` }}>
                                                            {g.monthLabel}
                                                        </div>
                                                    ));
                                                })()}
                                            </div>
                                        )}
                                        <div className={`flex w-full ${scheduleZoom === 'week' ? 'h-5' : 'h-10'}`}>
                                            {timeUnits.map((u, i) => (
                                                <div key={i} className={`flex items-center justify-center border-r border-slate-200 dark:border-slate-800/80 last:border-r-0 text-slate-500 font-bold tracking-wider ${scheduleZoom === 'week' ? 'text-[9px]' : 'text-[10px]'} overflow-hidden whitespace-nowrap text-ellipsis px-0.5`} style={{ width: `${u.widthPercent}%` }}>
                                                    {u.label}
                                                </div>
                                            ))}
                                        </div>
                                    </div>

                                    <div className="flex-1 relative">
                                        <div className="absolute top-0 bottom-0 left-0 right-0 flex pointer-events-none z-0">
                                            {timeUnits.map((u, i) => (
                                                <div key={`grid-${i}`} className="border-r border-slate-200 dark:border-slate-800/50 border-dashed shrink-0" style={{ width: `${u.widthPercent}%` }}></div>
                                            ))}
                                        </div>

                                        <div className="absolute top-0 left-0 right-0 flex flex-col z-10 py-[1px]">
                                            {displayedProjects.map((p) => {
                                                const pStart = new Date(p.spmk);
                                                const pEnd = new Date(p.deadline);

                                                const leftPercent = getDatePercent(pStart);
                                                const widthPercent = getDatePercent(pEnd) - leftPercent;

                                                const progress = Number(p.progress || 0);

                                                let pStatus = "On Progress";
                                                if (p.isPending) {
                                                    pStatus = "Pending";
                                                } else if (progress >= 100) {
                                                    pStatus = "Done";
                                                } else if (p.deadline) {
                                                    pStatus = getMicroStatus(progress, p.deadline);
                                                }

                                                let bgColor = '#3b82f6'; // blue-500
                                                let textColor = 'text-white';
                                                if (pStatus === "Pending") bgColor = '#64748b'; // slate-500
                                                else if (pStatus === "Done") bgColor = '#10b981'; // emerald-500
                                                else if (pStatus === "Terlambat") bgColor = '#ef4444'; // red-500
                                                else if (pStatus === "Beresiko") {
                                                    bgColor = '#facc15'; // yellow-400
                                                    textColor = 'text-slate-800';
                                                }

                                                return (
                                                    <div key={`bar-${p.id}`} className="h-14 flex items-center relative">
                                                        <div className="absolute h-8 rounded-full shadow-sm flex items-center transition-all duration-300 overflow-visible min-w-[20px] group cursor-pointer hover:brightness-110 hover:shadow-md z-10 hover:z-50"
                                                            onClick={() => { setActiveScheduleProject(p); setActiveTab('schedule'); }}
                                                            style={{
                                                                left: `${Math.max(0, leftPercent)}%`,
                                                                width: `${Math.min(100 - Math.max(0, leftPercent), widthPercent)}%`,
                                                                backgroundColor: bgColor
                                                            }}
                                                        >
                                                            <div className="absolute top-0 bottom-0 left-0 bg-white/20 rounded-l-full" style={{ width: `${progress}%` }}></div>
                                                            <span className={`relative z-10 px-3 text-xs font-bold truncate drop-shadow-md ${textColor}`}>{progress}%</span>

                                                            {/* Hover Tooltip */}
                                                            <div className="absolute top-full left-1/2 -translate-x-1/2 mt-2 opacity-0 invisible group-hover:opacity-100 group-hover:visible transition-all duration-200 z-50 bg-slate-800 text-white px-3 py-2 rounded-lg shadow-xl text-[10px] whitespace-nowrap flex flex-col gap-1 border border-slate-700/50 pointer-events-none">
                                                                <div className="absolute -top-1.5 left-1/2 -translate-x-1/2 border-l-4 border-r-4 border-b-4 border-transparent border-b-slate-800"></div>
                                                                <div className="font-bold text-slate-300 border-b border-slate-600 pb-1 mb-0.5">{p.name}</div>
                                                                {p.isPending && (
                                                                    <div className="mb-1 pb-1 border-b border-slate-600/50">
                                                                        <div className="text-orange-400 font-bold mb-0.5">Status: PENDING</div>
                                                                        <div className="text-slate-400 whitespace-normal w-48 leading-tight italic">"{p.pendingReason}"</div>
                                                                        {p.pendingDate && <div className="text-[8px] text-slate-500 mt-1">Sejak: {p.pendingDate}</div>}
                                                                    </div>
                                                                )}
                                                                <div className="flex items-center justify-between gap-4">
                                                                    <span className="text-slate-400">SPMK:</span>
                                                                    <span className="font-semibold text-emerald-400">{formatDateIndo(p.spmk)}</span>
                                                                </div>
                                                                <div className="flex items-center justify-between gap-4">
                                                                    <span className="text-slate-400">Akhir Kontrak:</span>
                                                                    <span className="font-semibold text-rose-400">{formatDateIndo(p.deadline)}</span>
                                                                </div>
                                                            </div>
                                                        </div>
                                                    </div>
                                                )
                                            })}
                                        </div>

                                        {(() => {
                                            const today = new Date();
                                            if (today >= chartStart && today <= chartEnd) {
                                                const todayOffset = (today - chartStart) / (1000 * 60 * 60 * 24);
                                                const todayPercent = (todayOffset / totalDays) * 100;
                                                return (
                                                    <div
                                                        className="absolute top-0 bottom-0 border-l-2 border-rose-500 border-dashed z-20 pointer-events-none"
                                                        style={{ left: `${todayPercent}%` }}
                                                    >
                                                        <div className="bg-rose-500 text-white text-[9px] font-bold px-2 py-0.5 rounded-full absolute top-2 -translate-x-1/2 shadow-[0_0_15px_rgba(244,63,94,0.6)] animate-pulse whitespace-nowrap flex items-center gap-1">
                                                            <div className="w-1.5 h-1.5 bg-white rounded-full"></div> Today
                                                        </div>
                                                    </div>
                                                );
                                            }
                                            return null;
                                        })()}
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                );
            };

            const renderPenugasan = () => {
                const filteredAssignments = assignments.filter(asg => {
                    // Filter pekerjaan yang sudah selesai (akhir kontrak < hari ini) agar tampilan tetap clean
                    if (asg.endDate) {
                        const today = new Date();
                        today.setHours(0, 0, 0, 0);
                        const end = new Date(asg.endDate);
                        end.setHours(0, 0, 0, 0);

                        if (end < today) {
                            return false;
                        }
                    }

                    const search = searchAssignmentTab.toLowerCase();
                    const matchJob = (asg.jobName || '').toLowerCase().includes(search);
                    const matchLpse = (asg.lpseName || '').toLowerCase().includes(search);

                    const matchExpert = (asg.experts || []).some(expPlot => {
                        const exp = experts.find(e => e.id === expPlot.expertId);
                        const expertName = exp ? exp.name.toLowerCase() : '';
                        return expertName.includes(search) || (expPlot.certificateName || '').toLowerCase().includes(search) || (expPlot.additionalCertificates || []).some(c => (c || '').toLowerCase().includes(search));
                    });

                    return matchJob || matchLpse || matchExpert;
                });

                return (
                    <div className="space-y-6 fade-in pb-12">
                        <div className="bg-white/60 dark:bg-slate-900/60 backdrop-blur-xl rounded-2xl shadow-sm border border-white/60 dark:border-slate-800 p-4 sm:p-6 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 transition-colors">
                            <div>
                                <h3 className="text-xl font-bold text-slate-800 dark:text-slate-100 flex items-center gap-2">
                                    <div className="p-2 bg-indigo-100 dark:bg-indigo-900/50 text-indigo-600 dark:text-indigo-400 rounded-lg">
                                        <Icon name="briefcase" size={24} />
                                    </div>
                                    Penugasan Tenaga Ahli
                                </h3>
                                <p className="text-sm text-slate-500 mt-1">Kelola plotting penugasan tenaga ahli untuk proyek.</p>
                            </div>
                            <div className="flex gap-2 w-full sm:w-auto flex-wrap sm:flex-nowrap">
                                <div className="relative flex-1 sm:w-64 min-w-[200px]">
                                    <input type="text" placeholder="Cari penugasan..." value={searchAssignmentTab} onChange={(e) => setSearchAssignmentTab(e.target.value)} className="w-full pl-10 pr-4 py-2 border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-800 dark:text-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500 text-sm transition-shadow" />
                                    <Icon name="search" size={16} className="absolute left-3 top-3 text-slate-400" />
                                </div>
                                <button onClick={() => {
                                    setPrintData({ type: 'expert_assignment' });
                                }} className="px-4 py-2 bg-white dark:bg-slate-800 text-slate-800 dark:text-slate-200 rounded-xl shadow-sm hover:bg-slate-50 dark:hover:bg-slate-700 transition-all font-semibold flex items-center gap-2 text-sm shrink-0 border border-slate-300 dark:border-slate-700">
                                    <Icon name="printer" size={18} /> Ekspor Laporan
                                </button>
                                <button onClick={() => setModalConfig({ isOpen: true, type: 'assignment', mode: 'add', data: null })} className="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl shadow-md transition-all font-semibold flex items-center gap-2 text-sm shrink-0">
                                    <Icon name="plus" size={18} /> Tambah Penugasan
                                </button>
                            </div>
                        </div>

                        <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-6">
                            {filteredAssignments.map(asg => {
                                return (
                                    <div key={asg.id} className="bg-white dark:bg-slate-900 rounded-2xl shadow-sm hover:shadow-md border border-slate-200 dark:border-slate-800 overflow-hidden transition-all group flex flex-col">
                                        <div className="p-5 border-b border-slate-100 dark:border-slate-800 flex justify-between items-start bg-slate-50/50 dark:bg-slate-900/50">
                                            <div>
                                                <h4 className="font-bold text-slate-800 dark:text-slate-100 text-lg leading-tight mb-2">{asg.jobName}</h4>
                                                <div className="flex flex-col items-start gap-2">
                                                    <span className="text-[10px] px-2 py-0.5 rounded-full font-bold bg-indigo-100 text-indigo-700 dark:bg-indigo-900/30 dark:text-indigo-400 border border-indigo-200 dark:border-indigo-800">
                                                        {asg.tenderType}
                                                    </span>
                                                    <p className="text-xs font-semibold text-slate-600 dark:text-slate-400">{asg.lpseName}</p>
                                                </div>
                                            </div>
                                            <div className="flex gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                                                <button onClick={() => setModalConfig({ isOpen: true, type: 'assignment', mode: 'edit', data: asg })} className="p-1.5 text-blue-600 hover:bg-blue-50 dark:text-blue-400 dark:hover:bg-slate-800 rounded-lg" title="Edit Penugasan"><Icon name="edit-3" size={16} /></button>
                                                <button onClick={() => {
                                                    if (confirm(`Hapus pekerjaan ${asg.jobName}?`)) handleAssignmentAction('delete', { id: asg.id });
                                                }} className="p-1.5 text-red-600 hover:bg-red-50 dark:text-red-400 dark:hover:bg-slate-800 rounded-lg" title="Hapus"><Icon name="trash-2" size={16} /></button>
                                            </div>
                                        </div>

                                        <div className="p-5 space-y-3 flex-1">
                                            <div className="flex flex-col gap-3 mb-3">
                                                <div>
                                                    <p className="text-xs text-slate-500 uppercase tracking-wider font-bold mb-0.5">Tipe Proyek & Kontrak</p>
                                                    <p className="text-sm text-slate-700 dark:text-slate-300 font-medium">{asg.projectType || 'Pengawasan'} - {asg.contractType}</p>
                                                </div>
                                                <div>
                                                    <p className="text-[10px] text-slate-500 font-bold mb-0.5">SPMK - Berakhir - Durasi</p>
                                                    <p className="text-xs font-semibold text-slate-700 dark:text-slate-300">
                                                        {asg.startDate ? formatDateIndo(asg.startDate) : '-'} s/d {asg.endDate ? formatDateIndo(asg.endDate) : '-'} ({asg.duration || 0} Hari)
                                                    </p>
                                                </div>
                                            </div>

                                            <div className="pt-3 border-t border-slate-100 dark:border-slate-800">
                                                <p className="text-xs text-slate-500 uppercase tracking-wider font-bold mb-2 flex items-center gap-1.5"><Icon name="users" size={12} /> Tenaga Ahli Diplot ({(asg.experts || []).length})</p>
                                                {(asg.experts || []).length === 0 ? (
                                                    <p className="text-xs text-slate-400 italic">Belum ada tenaga ahli.</p>
                                                ) : (
                                                    <div className="space-y-2">
                                                        {(asg.experts || []).map((expPlot, idx) => {
                                                            const exp = experts.find(e => e.id === expPlot.expertId);
                                                            const expertName = exp ? exp.name : 'Unknown';
                                                            const certObj = exp ? (exp.certificates || []).find(c => c.certName === expPlot.certificateName) : null;
                                                            let certDisplay = certObj && certObj.certLevel ? `${expPlot.certificateName} (${certObj.certLevel})` : expPlot.certificateName;
                                                            (expPlot.additionalCertificates || []).forEach(addCert => {
                                                                if (!addCert) return;
                                                                const cObj = exp ? (exp.certificates || []).find(c => c.certName === addCert) : null;
                                                                const cDisplay = cObj && cObj.certLevel ? `${addCert} (${cObj.certLevel})` : addCert;
                                                                if (!certDisplay) certDisplay = cDisplay;
                                                                else certDisplay += ` & ${cDisplay}`;
                                                            });
                                                            return (
                                                                <div key={idx} className="flex items-center justify-between bg-slate-50 dark:bg-slate-800/50 p-2 rounded-lg border border-slate-100 dark:border-slate-700/50">
                                                                    <div>
                                                                        <p className="text-xs font-bold text-slate-700 dark:text-slate-200">
                                                                            {expertName}
                                                                            {expPlot.role && <span className="font-normal text-slate-500">{" "}({expPlot.role})</span>}
                                                                        </p>
                                                                        {(expPlot.certificateName || (expPlot.additionalCertificates || []).length > 0) && <p className="text-[10px] text-slate-500">SKA: {certDisplay}</p>}
                                                                    </div>
                                                                    <div className="text-right">
                                                                        <span className="text-[10px] font-bold bg-indigo-100 text-indigo-700 dark:bg-indigo-900/50 dark:text-indigo-300 px-2 py-0.5 rounded-full">{expPlot.manMonth} MM</span>
                                                                    </div>
                                                                </div>
                                                            )
                                                        })}
                                                    </div>
                                                )}
                                            </div>
                                        </div>
                                    </div>
                                );
                            })}

                            {filteredAssignments.length === 0 && (
                                <div className="col-span-full py-12 text-center text-slate-400">
                                    <Icon name="briefcase" size={48} className="mx-auto mb-3 opacity-20" />
                                    <p className="text-lg font-medium text-slate-500">Belum ada penugasan tenaga ahli.</p>
                                    <p className="text-sm mt-1">Klik "Tambah Penugasan" untuk membuat data baru.</p>
                                </div>
                            )}
                        </div>
                    </div>
                );
            };

            const renderTenagaAhli = () => {
                const filteredExperts = experts.filter(e => {
                    const search = searchExpertTab.toLowerCase();
                    const matchName = e.name?.toLowerCase().includes(search);
                    const matchBidangIlmu = e.bidangIlmu?.toLowerCase().includes(search);
                    const matchCert = e.certificates?.some(c => c.certName?.toLowerCase().includes(search));
                    return matchName || matchBidangIlmu || matchCert;
                });

                const assignedExpertIds = new Set();
                assignments.forEach(asg => {
                    if (asg.experts) {
                        asg.experts.forEach(ex => assignedExpertIds.add(ex.expertId));
                    }
                });
                const assignedCount = experts.filter(e => assignedExpertIds.has(e.id)).length;
                const unassignedCount = experts.length - assignedCount;

                return (
                    <div className="space-y-6 fade-in pb-12">
                        <div className="bg-white/60 dark:bg-slate-900/60 backdrop-blur-xl rounded-2xl shadow-sm border border-white/60 dark:border-slate-800 p-4 sm:p-6 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 transition-colors">
                            <div>
                                <h3 className="text-xl font-bold text-slate-800 dark:text-slate-100 flex items-center gap-2">
                                    <div className="p-2 bg-indigo-100 dark:bg-indigo-900/50 text-indigo-600 dark:text-indigo-400 rounded-lg">
                                        <Icon name="award" size={24} />
                                    </div>
                                    Tenaga Ahli & Tender LPSE
                                </h3>
                                <p className="text-sm text-slate-500 mt-1">Kelola data tenaga ahli, sertifikat keahlian, dan alokasi tender LPSE.</p>
                            </div>
                            <div className="flex gap-2 w-full sm:w-auto">
                                <div className="relative flex-1 sm:w-64">
                                    <input type="text" placeholder="Cari tenaga ahli..." value={searchExpertTab} onChange={(e) => { setSearchExpertTab(e.target.value); setExpertPage(1); }} className="w-full pl-10 pr-4 py-2 border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-800 dark:text-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500 text-sm transition-shadow" />
                                    <Icon name="search" size={16} className="absolute left-3 top-3 text-slate-400" />
                                </div>
                                <button onClick={() => setModalConfig({ isOpen: true, type: 'expert', mode: 'add', data: null })} className="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl shadow-md transition-all font-semibold flex items-center gap-2 text-sm shrink-0">
                                    <Icon name="plus" size={18} /> Tambah
                                </button>
                                <button onClick={() => setModalConfig({ isOpen: true, type: 'import_expert', mode: 'add', data: null })} className="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl shadow-md transition-all font-semibold flex items-center gap-2 text-sm shrink-0">
                                    <Icon name="file-text" size={18} /> Import Excel
                                </button>
                            </div>
                        </div>

                        <div className="grid grid-cols-1 md:grid-cols-3 gap-3 sm:gap-4">
                            <MetricCard title="Total Tenaga Ahli" value={experts.length} trend="Terdaftar" trendUp={true} icon={<div className="text-blue-600"><Icon name="users" size={24} /></div>} color="bg-blue-50" />
                            <MetricCard title="Aktif Penugasan" value={assignedCount} trend="Ditugaskan LPSE" trendUp={true} icon={<div className="text-emerald-600"><Icon name="briefcase" size={24} /></div>} color="bg-emerald-50" />
                            <MetricCard title="Belum Ditugaskan" value={unassignedCount} trend="Tersedia" trendUp={true} icon={<div className="text-amber-600"><Icon name="clock" size={24} /></div>} color="bg-amber-50" />
                        </div>

                        <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-6">
                            {(() => {
                                const ITEMS_PER_PAGE = 30;
                                const totalPages = Math.ceil(filteredExperts.length / ITEMS_PER_PAGE);
                                const currentExperts = filteredExperts.slice((expertPage - 1) * ITEMS_PER_PAGE, expertPage * ITEMS_PER_PAGE);

                                return (
                                    <>
                                        {currentExperts.map(exp => (
                                            <div key={exp.id} className="bg-white dark:bg-slate-900 rounded-2xl shadow-sm hover:shadow-md border border-slate-200 dark:border-slate-800 overflow-hidden transition-all group">
                                    <div className="p-5 border-b border-slate-100 dark:border-slate-800 flex justify-between items-start bg-slate-50/50 dark:bg-slate-900/50">
                                        <div>
                                            <div className="flex items-center gap-2 mb-1">
                                                <h4 className="font-bold text-slate-800 dark:text-slate-100 text-lg">{exp.name}</h4>
                                            </div>
                                            <div className="flex flex-col gap-1 mt-2">
                                                <p className="text-xs text-slate-600 dark:text-slate-400 flex items-center gap-1.5"><Icon name="phone" size={12} /> {exp.phone || '-'}</p>
                                                {exp.bidangIlmu && (
                                                    <p className="text-xs text-slate-600 dark:text-slate-400 flex items-center gap-1.5" title="Jenjang & Bidang Ilmu">
                                                        <Icon name="graduation-cap" size={12} /> {exp.bidangIlmu}
                                                    </p>
                                                )}
                                                {exp.perusahaan && (
                                                    <p className="text-xs text-slate-600 dark:text-slate-400 flex items-center gap-1.5" title="Perusahaan / Instansi">
                                                        <Icon name="building-2" size={12} /> {exp.perusahaan}
                                                    </p>
                                                )}
                                            </div>
                                        </div>
                                        <div className="flex gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                                            <button onClick={() => setModalConfig({ isOpen: true, type: 'expert', mode: 'edit', data: exp })} className="p-1.5 text-blue-600 hover:bg-blue-50 dark:text-blue-400 dark:hover:bg-slate-800 rounded-lg" title="Edit Identitas"><Icon name="edit-3" size={16} /></button>
                                            <button onClick={() => {
                                                if (confirm(`Hapus tenaga ahli ${exp.name}?`)) handleExpertAction('delete', { id: exp.id });
                                            }} className="p-1.5 text-red-600 hover:bg-red-50 dark:text-red-400 dark:hover:bg-slate-800 rounded-lg" title="Hapus"><Icon name="trash-2" size={16} /></button>
                                        </div>
                                    </div>

                                    <div className="p-5 space-y-5">
                                        {/* Sertifikat Section */}
                                        <div>
                                            <div className="flex justify-between items-center mb-3">
                                                <h5 className="text-xs font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1.5">
                                                    <Icon name="shield" size={14} /> Sertifikat Keahlian
                                                </h5>
                                                <button onClick={() => setModalConfig({ isOpen: true, type: 'expert_cert', mode: 'add', data: { expertId: exp.id } })} className="text-[10px] font-bold text-indigo-600 hover:text-indigo-800 dark:text-indigo-400 bg-indigo-50 hover:bg-indigo-100 dark:bg-indigo-900/30 px-2 py-1 rounded-md transition-colors flex items-center gap-1">
                                                    <Icon name="plus" size={10} /> Tambah
                                                </button>
                                            </div>
                                            <div className="space-y-2">
                                                {(!exp.certificates || exp.certificates.length === 0) ? (
                                                    <p className="text-xs text-slate-400 italic">Belum ada sertifikat.</p>
                                                ) : (
                                                    exp.certificates.map((cert, idx) => {
                                                        const isExpired = new Date(cert.expiredDate) < new Date();
                                                        return (
                                                            <div key={idx} className="bg-slate-50 dark:bg-slate-800/50 p-2.5 rounded-xl border border-slate-200 dark:border-slate-700 flex justify-between items-center group/cert">
                                                                <div>
                                                                    <div className="text-sm font-semibold text-slate-700 dark:text-slate-200">{cert.certName} <span className="text-xs font-normal text-slate-500">({cert.certLevel})</span></div>
                                                                    <div className={`text-[10px] font-bold mt-0.5 ${isExpired ? 'text-red-500' : 'text-emerald-500'}`}>
                                                                        Exp: {formatDateIndo(cert.expiredDate)} {isExpired && '(EXPIRED)'}
                                                                    </div>
                                                                </div>
                                                                <div className="flex gap-1">
                                                                    <button onClick={() => setModalConfig({ isOpen: true, type: 'expert_cert', mode: 'edit', data: { expertId: exp.id, certIndex: idx, cert } })} className="text-slate-400 hover:text-blue-600 p-1 rounded transition-colors opacity-0 group-hover/cert:opacity-100"><Icon name="edit-2" size={14} /></button>
                                                                    <button onClick={() => {
                                                                        if (confirm('Hapus sertifikat ini?')) {
                                                                            let updatedCerts = [...exp.certificates];
                                                                            updatedCerts.splice(idx, 1);
                                                                            handleExpertAction('update_certificates', { ...exp, certificates: updatedCerts });
                                                                        }
                                                                    }} className="text-slate-400 hover:text-red-600 p-1 rounded transition-colors opacity-0 group-hover/cert:opacity-100"><Icon name="trash-2" size={14} /></button>
                                                                </div>
                                                            </div>
                                                        );
                                                    })
                                                )}
                                            </div>
                                        </div>

                                        {/* Tender LPSE Section (Linked to Assignments) */}
                                        <div>
                                            <div className="flex justify-between items-center mb-3">
                                                <h5 className="text-xs font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1.5">
                                                    <Icon name="briefcase" size={14} /> Riwayat Tender LPSE
                                                </h5>
                                                {/* Tombol Tambah dihapus karena sudah terhubung ke menu Penugasan Tenaga Ahli */}
                                            </div>
                                            <div className="space-y-2">
                                                {(() => {
                                                    const expertAssignments = [];
                                                    assignments.forEach(asg => {
                                                        // Filter out expired assignments
                                                        if (asg.endDate) {
                                                            const today = new Date();
                                                            today.setHours(0, 0, 0, 0);
                                                            const end = new Date(asg.endDate);
                                                            end.setHours(0, 0, 0, 0);

                                                            if (end < today) {
                                                                return;
                                                            }
                                                        }

                                                        const plot = (asg.experts || []).find(e => e.expertId === exp.id);
                                                        if (plot) {
                                                            expertAssignments.push({ ...asg, plotData: plot });
                                                        }
                                                    });

                                                    if (expertAssignments.length === 0) {
                                                        return <p className="text-xs text-slate-400 italic">Belum ada penugasan terdaftar.</p>;
                                                    }
                                                    return expertAssignments.map((tender, idx) => (
                                                        <div key={idx} className="bg-slate-50 dark:bg-slate-800/50 p-2.5 rounded-xl border border-slate-200 dark:border-slate-700 flex justify-between items-center group/tender">
                                                            <div>
                                                                <div className="text-xs font-semibold text-slate-700 dark:text-slate-200">{tender.lpseName}</div>
                                                                <div className={`text-[10px] font-bold px-1.5 py-0.5 rounded inline-block mt-1 ${tender.contractType === 'Waktu Penugasan' ? 'bg-orange-100 text-orange-700 border-orange-200 dark:bg-orange-900/50 dark:text-orange-300 dark:border-orange-800' : 'bg-slate-200 text-slate-600 border-slate-300 dark:bg-slate-700 dark:text-slate-300 dark:border-slate-600'}`}>
                                                                    {tender.jobName} ({tender.tenderType})
                                                                </div>
                                                                <div className="mt-1.5 text-[10px] text-slate-500 dark:text-slate-400 font-semibold flex flex-wrap gap-x-3 gap-y-1">
                                                                    {
                                                                        (() => {
                                                                            const certObj = exp.certificates?.find(c => c.certName === tender.plotData.certificateName);
                                                                            let certDisplay = certObj && certObj.certLevel ? `${tender.plotData.certificateName} (${certObj.certLevel})` : (tender.plotData.certificateName || '-');
                                                                            (tender.plotData.additionalCertificates || []).forEach(addCert => {
                                                                                if (!addCert) return;
                                                                                const cObj = exp.certificates?.find(c => c.certName === addCert);
                                                                                const cDisplay = cObj && cObj.certLevel ? `${addCert} (${cObj.certLevel})` : addCert;
                                                                                if (certDisplay === '-') certDisplay = cDisplay;
                                                                                else certDisplay += ` & ${cDisplay}`;
                                                                            });
                                                                            return <span>SKA: {certDisplay}</span>;
                                                                        })()
                                                                    }
                                                                    <span>MM: {tender.plotData.manMonth}</span>
                                                                    <span>Kontrak: {tender.contractType || '-'}</span>
                                                                    <span>Selesai: {tender.endDate ? new Date(tender.endDate).toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' }) : '-'}</span>
                                                                </div>
                                                            </div>
                                                        </div>
                                                    ));
                                                })()}
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                        ))}
                                        {filteredExperts.length === 0 && (
                                            <div className="col-span-full py-12 text-center text-slate-500 bg-white/50 dark:bg-slate-900/50 rounded-2xl border border-dashed border-slate-300 dark:border-slate-700">
                                                <Icon name="award" size={48} className="mx-auto mb-3 opacity-20" />
                                                <p>Tidak ada data tenaga ahli ditemukan.</p>
                                            </div>
                                        )}
                                        {totalPages > 1 && (
                                            <div className="col-span-full flex justify-center items-center mt-6 space-x-3 fade-in">
                                                <button
                                                    onClick={() => setExpertPage(prev => Math.max(prev - 1, 1))}
                                                    disabled={expertPage === 1}
                                                    className="px-4 py-2 rounded-xl border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-700 dark:text-slate-300 disabled:opacity-30 hover:bg-slate-100 dark:hover:bg-slate-700 transition-all font-semibold flex items-center gap-2 shadow-sm"
                                                >
                                                    <Icon name="chevron-left" size={16} /> Sebelumnya
                                                </button>
                                                <span className="text-slate-600 dark:text-slate-400 font-medium px-4 py-2 bg-slate-100 dark:bg-slate-800 rounded-xl text-sm border border-slate-200 dark:border-slate-700">
                                                    Hal {expertPage} dari {totalPages}
                                                </span>
                                                <button
                                                    onClick={() => setExpertPage(prev => Math.min(prev + 1, totalPages))}
                                                    disabled={expertPage === totalPages}
                                                    className="px-4 py-2 rounded-xl border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-700 dark:text-slate-300 disabled:opacity-30 hover:bg-slate-100 dark:hover:bg-slate-700 transition-all font-semibold flex items-center gap-2 shadow-sm"
                                                >
                                                    Selanjutnya <Icon name="chevron-right" size={16} />
                                                </button>
                                            </div>
                                        )}
                                    </>
                                );
                            })()}
                        </div>
                    </div>
                );
            };

            return (
                <>
                    {renderPrintExecutiveReport()}
                    <div id="main-ui-wrapper" className="flex h-screen text-slate-800 dark:text-slate-200 bg-transparent transition-colors duration-200 relative overflow-hidden print:hidden">
                        {/* Subtle Glowing Orbs for Glassmorphism Depth */}
                        <div className="absolute inset-0 pointer-events-none z-0 overflow-hidden">
                            <div className="absolute -top-[20%] -left-[10%] w-[60vw] h-[60vw] rounded-full bg-indigo-500/10 dark:bg-indigo-600/10 blur-[100px] mix-blend-multiply dark:mix-blend-screen opacity-60"></div>
                            <div className="absolute top-[40%] -right-[10%] w-[50vw] h-[50vw] rounded-full bg-fuchsia-500/10 dark:bg-fuchsia-600/10 blur-[120px] mix-blend-multiply dark:mix-blend-screen opacity-60"></div>
                        </div>

                        <div className="relative z-10 flex w-full h-full">
                            {renderFuturisticConfirm()}
                            {renderDeveloperPromptModal()}
                            {renderDominoModal()}

                            {renderKPIInfoModal()}
                            {showProjectTypeModal && (
                                <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-sm z-50 flex items-center justify-center p-4">
                                    <div className="glass-card p-6 rounded-3xl shadow-2xl max-w-sm w-full animate-in fade-in zoom-in duration-200">
                                        <div className="flex items-center justify-between mb-5">
                                            <h3 className="text-lg font-bold text-slate-800 dark:text-slate-200 flex items-center gap-2">
                                                <Icon name="briefcase" size={20} className="text-blue-500" /> Detail Tipe Proyek
                                            </h3>
                                            <button onClick={() => setShowProjectTypeModal(false)} className="text-slate-400 hover:text-slate-600 dark:hover:text-slate-200"><Icon name="x" size={20} /></button>
                                        </div>
                                        <div className="space-y-3">
                                            {Object.entries(
                                                computedProjects.reduce((acc, p) => {
                                                    let type = p.type || "Lainnya";
                                                    if (type.trim().toLowerCase() === "perencana") type = "Perencanaan";
                                                    if (type.trim().toLowerCase() === "pengawas") type = "Pengawasan";

                                                    acc[type] = (acc[type] || 0) + 1;
                                                    return acc;
                                                }, {})
                                            ).sort((a, b) => b[1] - a[1]).map(([type, count]) => (
                                                <div key={type} className="flex justify-between items-center p-3 bg-slate-50 dark:bg-slate-900/50 rounded-lg border border-slate-100 dark:border-slate-700">
                                                    <span className="font-semibold text-sm text-slate-700 dark:text-slate-300">{type}</span>
                                                    <span className="font-bold text-blue-600 dark:text-blue-400 bg-blue-100 dark:bg-blue-900/30 px-2.5 py-1 rounded-full text-xs">{count} Proyek</span>
                                                </div>
                                            ))}
                                            <div className="flex justify-between items-center p-3 mt-4 border-t border-slate-200 dark:border-slate-700">
                                                <span className="font-black text-sm text-slate-800 dark:text-slate-200">Total Keseluruhan</span>
                                                <span className="font-black text-slate-800 dark:text-slate-200">{computedProjects.length} Proyek</span>
                                            </div>
                                        </div>
                                        <div className="mt-6 flex justify-end">
                                            <button onClick={() => setShowProjectTypeModal(false)} className="px-4 py-2 w-full rounded-lg text-sm font-medium text-white bg-blue-600 hover:bg-blue-700 transition-colors shadow-md">Tutup</button>
                                        </div>
                                    </div>
                                </div>
                            )}
                            {showPrintModal && (
                                <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-sm z-50 flex items-center justify-center p-4">
                                    <div className="glass-card p-6 rounded-3xl shadow-2xl max-w-md w-full animate-in fade-in zoom-in duration-200">
                                        <div className="flex items-center justify-between mb-5">
                                            <h3 className="text-lg font-bold text-slate-800 dark:text-slate-200">Opsi Ekspor Laporan</h3>
                                            <button onClick={() => setShowPrintModal(false)} className="text-slate-400 hover:text-slate-600 dark:hover:text-slate-200"><Icon name="x" size={20} /></button>
                                        </div>
                                        <div className="space-y-4">
                                            <div>
                                                <label className="block text-sm font-semibold text-slate-700 dark:text-slate-300 mb-1.5">Tipe Proyek</label>
                                                <select value={printOptions.projectType} onChange={(e) => setPrintOptions({ ...printOptions, projectType: e.target.value })} className="w-full bg-slate-50 dark:bg-slate-900/50 border border-slate-200 dark:border-slate-700 rounded-lg px-3 py-2 text-sm text-slate-700 dark:text-slate-300 focus:outline-none focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500">
                                                    <option value="Semua">Semua Tipe Proyek</option>
                                                    <option value="Perencanaan">Hanya Perencanaan</option>
                                                    <option value="Pengawasan">Hanya Pengawasan</option>
                                                    <option value="Manajemen Konstruksi">Hanya Manajemen Konstruksi</option>
                                                </select>
                                            </div>
                                            <div>
                                                <label className="block text-sm font-semibold text-slate-700 dark:text-slate-300 mb-1.5">Cakupan Laporan</label>
                                                <select value={printOptions.section} onChange={(e) => setPrintOptions({ ...printOptions, section: e.target.value })} className="w-full bg-slate-50 dark:bg-slate-900/50 border border-slate-200 dark:border-slate-700 rounded-lg px-3 py-2 text-sm text-slate-700 dark:text-slate-300 focus:outline-none focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500">
                                                    <option value="Semua">Lengkap (Status Proyek & Penugasan Pegawai)</option>
                                                    <option value="ProyekSaja">Hanya Laporan Status Proyek Saja</option>
                                                    <option value="PegawaiSaja">Hanya Rincian Penugasan Pegawai Saja</option>
                                                </select>
                                            </div>
                                        </div>
                                        <div className="mt-6 flex justify-end gap-3">
                                            <button onClick={() => setShowPrintModal(false)} className="px-4 py-2 rounded-lg text-sm font-medium text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors">Batal</button>
                                            <button onClick={() => {
                                                setShowPrintModal(false);
                                                handleExportExcel(printOptions);
                                            }} className="px-4 py-2 rounded-lg text-sm font-medium text-white bg-green-600 hover:bg-green-700 transition-colors flex items-center gap-2 shadow-md">
                                                <Icon name="file-text" size={16} /> Ekspor Excel
                                            </button>
                                            <button onClick={() => {
                                                setShowPrintModal(false);
                                                setPrintData({ type: 'custom', options: printOptions });
                                            }} className="px-4 py-2 rounded-lg text-sm font-medium text-white bg-indigo-600 hover:bg-indigo-700 transition-colors flex items-center gap-2 shadow-md">
                                                <Icon name="printer" size={16} /> Cetak PDF
                                            </button>
                                        </div>
                                    </div>
                                </div>
                            )}
                            <aside className="hidden lg:flex lg:relative z-auto w-[270px] my-5 ml-5 rounded-[2rem] glass-panel text-slate-700 dark:text-slate-200 flex-col shrink-0 transition-all duration-300 shadow-xl overflow-hidden">
                                <div className="p-7 pb-5 border-b border-slate-200/50 dark:border-slate-700/30">
                                    <div className="flex items-center gap-3 w-full mb-3">
                                        <div className="w-10 h-10 shrink-0 flex items-center justify-center">
                                            <img src={logoImg} alt="Logo SIDAMON" className="w-full h-full object-contain drop-shadow-sm" />
                                        </div>
                                        <div className="flex flex-col min-w-0">
                                            <h1 className="text-lg font-black tracking-widest leading-none text-slate-900 dark:text-white">SIDAMON</h1>
                                            <span className="text-xs font-bold tracking-wider text-slate-600 dark:text-slate-400 mt-1 truncate">Gaharu Sempana Group</span>
                                        </div>
                                    </div>
                                    <p className="text-[10px] text-slate-500 uppercase tracking-widest font-bold whitespace-nowrap">Sistem Database & Monitoring</p>
                                </div>
                                <nav className="flex-1 px-4 space-y-1 mt-4 overflow-y-auto pb-4">
                                    <div className="text-[10px] font-bold text-slate-400 dark:text-slate-500 uppercase tracking-widest mb-1 mt-2 first:mt-0 px-2">Overview</div>
                                    <SidebarItem icon={<Icon name="layout-dashboard" size={20} />} label="Dashboard" isActive={activeTab === 'dashboard'} onClick={() => handleTabChange('dashboard')} />
                                    
                                    <div className="text-[10px] font-bold text-slate-400 dark:text-slate-500 uppercase tracking-widest mb-1 mt-4 px-2">Project Management & Control</div>
                                    <SidebarItem icon={<Icon name="briefcase" size={20} />} label="List Proyek" isActive={activeTab === 'proyek'} onClick={() => handleTabChange('proyek')} />
                                    <SidebarItem icon={<Icon name="calendar" size={20} />} label="Master Schedule" isActive={activeTab === 'master-schedule'} onClick={() => handleTabChange('master-schedule')} />
                                    <SidebarItem icon={<Icon name="users" size={20} />} label="Alokasi Tim" isActive={activeTab === 'tim'} onClick={() => handleTabChange('tim')} />
                                    <SidebarItem icon={<Icon name="calendar-days" size={20} />} label="Plotting Jadwal" isActive={activeTab === 'gantt'} onClick={() => handleTabChange('gantt')} />
                                    
                                    <div className="text-[10px] font-bold text-slate-400 dark:text-slate-500 uppercase tracking-widest mb-1 mt-4 px-2">Database & Assignment Experts</div>
                                    <SidebarItem icon={<Icon name="award" size={20} />} label="Tenaga Ahli" isActive={activeTab === 'ahli'} onClick={() => handleTabChange('ahli')} />
                                    <SidebarItem icon={<Icon name="briefcase" size={20} />} label="Penugasan Tenaga Ahli" isActive={activeTab === 'penugasan'} onClick={() => handleTabChange('penugasan')} />
                                    
                                    <div className="text-[10px] font-bold text-slate-400 dark:text-slate-500 uppercase tracking-widest mb-1 mt-4 px-2">Assets</div>
                                    <SidebarItem icon={<Icon name="box" size={20} />} label="Logistik & Inventaris" isActive={activeTab === 'inventaris'} onClick={() => handleTabChange('inventaris')} />
                                    
                                    <div className="text-[10px] font-bold text-slate-400 dark:text-slate-500 uppercase tracking-widest mb-1 mt-4 px-2">Performance</div>
                                    <SidebarItem icon={<Icon name="bar-chart" size={20} />} label="KPI & Evaluasi" isActive={activeTab === 'kpi'} onClick={() => handleTabChange('kpi')} />
                                </nav>
                            </aside>

                            <main className="flex-1 flex flex-col overflow-hidden bg-transparent relative min-w-0 transition-colors duration-200">
                                <header className="h-[72px] mx-5 lg:mx-8 mt-5 lg:mt-5 rounded-2xl glass-panel flex items-center justify-between px-6 lg:px-8 z-10 shrink-0 gap-3 shadow-md transition-all duration-300">
                                    <div className="flex items-center gap-2 sm:gap-3 min-w-0">
                                        <div className="min-w-0">
                                            <h2 className="text-sm sm:text-base lg:text-xl font-bold text-slate-800 dark:text-slate-100 truncate transition-colors duration-200">
                                                {activeTab === 'dashboard' && 'Ringkasan Proyek & Personil'}
                                                {activeTab === 'master-schedule' && 'Master Schedule Proyek'}
                                                {activeTab === 'proyek' && 'List Proyek'}
                                                {activeTab === 'tim' && 'Alokasi Tim'}
                                                {activeTab === 'ahli' && 'Tenaga Ahli & Tender LPSE'}
                                                {activeTab === 'penugasan' && 'Penugasan Tenaga Ahli'}
                                                {activeTab === 'gantt' && 'Ploting Jadwal & Timeline'}
                                                {activeTab === 'timesheet' && 'Timesheet Harian'}
                                                {activeTab === 'inventaris' && 'Logistik & Inventaris Alat'}
                                                {activeTab === 'kpi' && 'Evaluasi Tim & KPI'}
                                                {activeTab === 'schedule' && activeScheduleProject && `Time Schedule: ${activeScheduleProject.name}`}
                                            </h2>
                                            <p className="text-[10px] sm:text-xs text-slate-500 mt-0.5 hidden sm:block">Aplikasi Manajemen Proyek & Personil Tim Teknis</p>
                                        </div>
                                    </div>
                                    <div className="flex items-center gap-3 sm:gap-4 shrink-0">
                                        <button
                                            onClick={() => setIsDarkMode(!isDarkMode)}
                                            className={`relative inline-flex h-8 w-[64px] shrink-0 items-center rounded-full transition-all duration-500 shadow-inner hover:scale-105 ${isDarkMode ? 'bg-white border-slate-200 dark:bg-slate-200' : 'bg-slate-800 border-slate-700'}`}
                                            title={isDarkMode ? 'Ubah ke Light Mode' : 'Ubah ke Dark Mode'}
                                        >
                                            <span
                                                className={`inline-flex h-6 w-6 transform items-center justify-center rounded-full transition-transform duration-500 ease-out z-10 ${isDarkMode ? 'translate-x-[34px] bg-slate-900 text-white shadow-md' : 'translate-x-1 bg-white text-slate-800 shadow-md'}`}
                                            >
                                                <Icon name={isDarkMode ? "moon" : "sun"} size={14} strokeWidth={2.5} />
                                            </span>
                                        </button>
                                        <div className={`hidden sm:flex items-center gap-2 px-2 sm:px-3 py-1 sm:py-1.5 rounded-full text-xs font-semibold shadow-sm ${isLive ? 'bg-emerald-100 text-emerald-700 border border-emerald-200 dark:bg-emerald-900/30 dark:text-emerald-400 dark:border-emerald-800' : 'bg-amber-100 text-amber-700 border border-amber-200 dark:bg-amber-900/30 dark:text-amber-400 dark:border-amber-800'}`}>
                                            <Icon name="database" size={14} /> {isLive ? 'Live' : 'Offline'}
                                        </div>
                                        <div className={`sm:hidden w-2.5 h-2.5 rounded-full shrink-0 ${isLive ? 'bg-emerald-500' : 'bg-amber-500'}`} title={isLive ? 'Connected' : 'Offline'} />
                                        <button onClick={() => setShowPrintModal(true)} className="hidden sm:flex items-center gap-2 px-3 py-1.5 text-xs font-bold text-slate-700 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 transition-colors shadow-sm dark:bg-slate-800 dark:border-slate-700 dark:text-slate-200 dark:hover:bg-slate-700">
                                            <Icon name="printer" size={14} /> Ekspor Laporan
                                        </button>
                                        <button onClick={fetchFromGoogleSheets} disabled={loading || !GOOGLE_SCRIPT_URL} className={`p-2 rounded-full transition-all shrink-0 ${loading ? 'animate-spin text-blue-500' : 'text-slate-400 hover:text-blue-600 hover:bg-blue-50 dark:hover:bg-slate-800 dark:text-slate-500 dark:hover:text-blue-400'}`} title="Refresh Database">
                                            <Icon name="refresh-ccw" size={20} />
                                        </button>
                                    </div>
                                </header>

                                <div className="flex-1 overflow-auto p-3 sm:p-5 lg:p-8 pb-24 lg:pb-8 relative">
                                    {loading && (
                                        <div className="fixed bottom-6 right-6 bg-white dark:bg-slate-800 shadow-xl shadow-blue-900/10 border border-slate-200 dark:border-slate-700 rounded-2xl p-4 z-50 flex items-center gap-4 fade-in">
                                            <div className="p-2.5 bg-blue-50 dark:bg-blue-900/30 rounded-xl">
                                                <Icon name="refresh-ccw" size={20} className="animate-spin text-blue-600 dark:text-blue-400" />
                                            </div>
                                            <div>
                                                <p className="font-bold text-sm text-slate-800 dark:text-slate-200">Sinkronisasi Database</p>
                                                <p className="text-xs text-slate-500 dark:text-slate-400 mt-0.5">Menyimpan perubahan...</p>
                                            </div>
                                        </div>
                                    )}

                                    <div className="max-w-7xl mx-auto">
                                        {activeTab === 'dashboard' && renderDashboard()}
                                        {activeTab === 'master-schedule' && renderMasterSchedule()}
                                        {activeTab === 'proyek' && renderProyek()}
                                        {activeTab === 'tim' && renderTim()}
                                        {activeTab === 'ahli' && renderTenagaAhli()}
                                        {activeTab === 'penugasan' && renderPenugasan()}
                                        {activeTab === 'gantt' && renderGantt()}
                                        {activeTab === 'timesheet' && renderTimesheet()}
                                        {activeTab === 'inventaris' && renderInventory()}
                                        {activeTab === 'kpi' && renderKPI()}
                                        {activeTab === 'schedule' && renderTimeSchedule()}
                                    </div>
                                </div>
                            </main>

                            {printZoomProject && (
                                <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[100] flex items-center justify-center fade-in">
                                    <div className="bg-white dark:bg-slate-900 rounded-2xl shadow-2xl w-full max-w-sm p-6 transform scale-in border border-slate-200 dark:border-slate-800">
                                        <div className="text-center mb-6">
                                            <div className="w-16 h-16 bg-indigo-100 dark:bg-indigo-900/30 text-indigo-600 dark:text-indigo-400 rounded-full flex items-center justify-center mx-auto mb-4">
                                                <Icon name="printer" size={32} />
                                            </div>
                                            <h3 className="text-xl font-bold text-slate-800 dark:text-slate-100">Pilih Format PDF</h3>
                                            <p className="text-sm text-slate-500 mt-2">Pilih tingkat detail kalender untuk laporan cetak Time Schedule proyek <strong>{printZoomProject.name}</strong></p>
                                        </div>
                                        <div className="flex flex-col gap-3">
                                            <button
                                                onClick={() => { setScheduleZoom('month'); setPrintData({ type: 'project', id: printZoomProject.id }); setPrintZoomProject(null); }}
                                                className="w-full flex items-center gap-3 p-4 rounded-xl border-2 border-slate-200 dark:border-slate-700 hover:border-indigo-500 dark:hover:border-indigo-500 hover:bg-indigo-50 dark:hover:bg-indigo-900/20 transition-all text-left group"
                                            >
                                                <div className="w-10 h-10 rounded-lg bg-slate-100 dark:bg-slate-800 flex items-center justify-center group-hover:bg-indigo-100 dark:group-hover:bg-indigo-900/50 text-slate-500 group-hover:text-indigo-600 transition-colors">
                                                    <Icon name="calendar" size={20} />
                                                </div>
                                                <div>
                                                    <div className="font-bold text-slate-700 dark:text-slate-200">Bulanan</div>
                                                    <div className="text-xs text-slate-500">Tampilan ringkas per bulan</div>
                                                </div>
                                            </button>
                                            <button
                                                onClick={() => { setScheduleZoom('week'); setPrintData({ type: 'project', id: printZoomProject.id }); setPrintZoomProject(null); }}
                                                className="w-full flex items-center gap-3 p-4 rounded-xl border-2 border-slate-200 dark:border-slate-700 hover:border-indigo-500 dark:hover:border-indigo-500 hover:bg-indigo-50 dark:hover:bg-indigo-900/20 transition-all text-left group"
                                            >
                                                <div className="w-10 h-10 rounded-lg bg-slate-100 dark:bg-slate-800 flex items-center justify-center group-hover:bg-indigo-100 dark:group-hover:bg-indigo-900/50 text-slate-500 group-hover:text-indigo-600 transition-colors">
                                                    <Icon name="calendar-days" size={20} />
                                                </div>
                                                <div>
                                                    <div className="font-bold text-slate-700 dark:text-slate-200">Mingguan</div>
                                                    <div className="text-xs text-slate-500">Tampilan detail tiap minggu</div>
                                                </div>
                                            </button>
                                        </div>
                                        <div className="mt-6">
                                            <button onClick={() => setPrintZoomProject(null)} className="w-full px-4 py-3 text-sm font-semibold text-slate-600 dark:text-slate-400 bg-slate-100 dark:bg-slate-800 hover:bg-slate-200 dark:hover:bg-slate-700 rounded-xl transition-colors">Batal</button>
                                        </div>
                                    </div>
                                </div>
                            )}

                            {showPendingModal && pendingProjectData && (
                                <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[100] flex items-center justify-center fade-in p-4">
                                    <div className="bg-white dark:bg-slate-900 rounded-2xl shadow-2xl w-full max-w-md p-6 transform scale-in border border-slate-200 dark:border-slate-800">
                                        <div className="flex justify-between items-center mb-4">
                                            <h3 className="text-lg font-bold text-slate-800 dark:text-slate-100 flex items-center gap-2">
                                                <Icon name="alert-triangle" className="text-orange-500" />
                                                Pending Proyek
                                            </h3>
                                            <button onClick={() => { setShowPendingModal(false); setPendingProjectData(null); setPendingReasonText(""); }} className="text-slate-400 hover:text-slate-600 transition-colors">
                                                <Icon name="x" size={24} />
                                            </button>
                                        </div>
                                        <div className="mb-4">
                                            <p className="text-sm text-slate-600 dark:text-slate-400 mb-2">
                                                Anda akan mengubah status proyek <strong className="text-slate-800 dark:text-slate-200">{pendingProjectData.name}</strong> menjadi Pending. Status ini akan membebastugaskan tim sementara waktu.
                                            </p>
                                            <label className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1">
                                                Alasan Pending <span className="text-red-500">*</span>
                                            </label>
                                            <textarea
                                                value={pendingReasonText}
                                                onChange={(e) => setPendingReasonText(e.target.value)}
                                                className="w-full border border-slate-300 dark:border-slate-700 rounded-lg p-3 text-sm bg-white dark:bg-slate-800 text-slate-800 dark:text-slate-200 focus:ring-2 focus:ring-orange-500 focus:border-orange-500"
                                                rows="3"
                                                placeholder="Contoh: Menunggu konfirmasi revisi RAB dari klien..."
                                            ></textarea>
                                        </div>
                                        <div className="flex justify-end gap-2 mt-6">
                                            <button onClick={() => { setShowPendingModal(false); setPendingProjectData(null); setPendingReasonText(""); }} className="px-4 py-2 text-sm font-semibold text-slate-700 dark:text-slate-300 bg-slate-100 dark:bg-slate-800 hover:bg-slate-200 dark:hover:bg-slate-700 rounded-lg">Batal</button>
                                            <button onClick={handleTogglePendingSubmit} disabled={!pendingReasonText.trim()} className="px-4 py-2 text-sm font-bold text-white bg-orange-600 hover:bg-orange-700 disabled:opacity-50 disabled:cursor-not-allowed rounded-lg flex items-center gap-2">
                                                Simpan Status
                                            </button>
                                        </div>
                                    </div>
                                </div>
                            )}

                            {showResumeModal && resumeProjectData && (
                                <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[100] flex items-center justify-center fade-in p-4">
                                    <div className="bg-white dark:bg-slate-900 rounded-2xl shadow-2xl w-full max-w-md p-6 transform scale-in border border-slate-200 dark:border-slate-800">
                                        <div className="flex justify-between items-center mb-4">
                                            <h3 className="text-lg font-bold text-slate-800 dark:text-slate-100 flex items-center gap-2">
                                                <Icon name="play" className="text-emerald-500" />
                                                Lanjutkan Proyek
                                            </h3>
                                            <button onClick={() => { setShowResumeModal(false); setResumeProjectData(null); }} className="text-slate-400 hover:text-slate-600 transition-colors">
                                                <Icon name="x" size={24} />
                                            </button>
                                        </div>
                                        <div className="mb-4">
                                            <p className="text-sm text-slate-600 dark:text-slate-400 mb-2">
                                                Apakah Anda yakin ingin melanjutkan proyek <strong className="text-slate-800 dark:text-slate-200">{resumeProjectData.name}</strong>?
                                            </p>
                                            <p className="text-sm text-slate-600 dark:text-slate-400">
                                                Status proyek ini akan kembali aktif dan beban kerja akan mulai diperhitungkan kembali ke anggota tim yang bertugas.
                                            </p>
                                        </div>
                                        <div className="flex justify-end gap-2 mt-6">
                                            <button onClick={() => { setShowResumeModal(false); setResumeProjectData(null); }} className="px-4 py-2 text-sm font-semibold text-slate-700 dark:text-slate-300 bg-slate-100 dark:bg-slate-800 hover:bg-slate-200 dark:hover:bg-slate-700 rounded-lg">Batal</button>
                                            <button onClick={handleConfirmResumeProject} className="px-4 py-2 text-sm font-bold text-white bg-emerald-600 hover:bg-emerald-700 rounded-lg flex items-center gap-2">
                                                Ya, Lanjutkan Proyek
                                            </button>
                                        </div>
                                    </div>
                                </div>
                            )}

                            <ModalForm />
                            <ExpertModalForm />
                            <ImportExcelModal />
                            <ExpertCertModalForm />
                            <ExpertTenderModalForm />
                            <AssignmentModalForm />

                            {/* Lpse Manager Modal Inline */}
                            {showLpseManager && (
                                <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[120] flex items-center justify-center fade-in p-4">
                                    <div className="glass-card rounded-3xl shadow-2xl w-full max-w-md flex flex-col max-h-[80vh] overflow-hidden transform scale-in border border-slate-200 dark:border-slate-800">
                                        <div className="flex justify-between items-center p-5 border-b border-slate-100 dark:border-slate-700/50 bg-slate-50/50 dark:bg-slate-900/50">
                                            <h3 className="text-xl font-bold text-slate-800 dark:text-slate-100 flex items-center gap-2"><Icon name="settings" size={20} className="text-slate-400" /> Kelola Daftar LPSE</h3>
                                            <button onClick={() => setShowLpseManager(false)} type="button" className="text-slate-400 hover:bg-slate-200 hover:text-slate-600 dark:hover:bg-slate-700 dark:hover:text-slate-300 p-2 rounded-xl transition-colors"><Icon name="x" size={20} /></button>
                                        </div>
                                        <div className="p-6 overflow-y-auto custom-scrollbar flex-1">
                                            <form onSubmit={(e) => {
                                                e.preventDefault();
                                                const input = e.target.elements.newlpse.value.trim();
                                                if (!input) return;
                                                if (lpseList.includes(input)) {
                                                    setAlertModal({ isOpen: true, title: 'Duplikasi LPSE', message: "LPSE sudah ada di dalam daftar!" });
                                                    return;
                                                }
                                                handleUpdateLpseList([...lpseList, input]);
                                                e.target.reset();
                                            }} className="flex gap-2 mb-4">
                                                <input type="text" name="newlpse" placeholder="Tambah instansi LPSE baru..." className="flex-1 p-2 border border-slate-300 dark:border-slate-700 rounded-lg outline-none focus:border-indigo-500 bg-slate-50 dark:bg-slate-800 text-sm dark:text-slate-200" />
                                                <button type="submit" className="px-3 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg text-sm font-semibold transition-colors flex items-center gap-1"><Icon name="plus" size={16} /> Tambah</button>
                                            </form>
                                            <div className="space-y-2">
                                                {lpseList.length === 0 ? <p className="text-sm text-slate-400 italic">Daftar kosong.</p> : null}
                                                {lpseList.map((lpse, idx) => (
                                                    <div key={idx} className="flex justify-between items-center p-2.5 bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg">
                                                        <span className="text-sm text-slate-700 dark:text-slate-300">{lpse}</span>
                                                        <button onClick={() => {
                                                            if (confirm(`Hapus ${lpse} dari daftar LPSE?`)) {
                                                                handleUpdateLpseList(lpseList.filter(item => item !== lpse));
                                                            }
                                                        }} className="text-slate-400 hover:text-red-500 transition-colors p-1" title="Hapus"><Icon name="trash-2" size={16} /></button>
                                                    </div>
                                                ))}
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            )}

                            {/* Cert Manager Modal Inline */}
                            {showCertManager && (
                                <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[130] flex items-center justify-center fade-in p-4">
                                    <div className="glass-card rounded-3xl shadow-2xl w-full max-w-md flex flex-col max-h-[80vh] overflow-hidden transform scale-in border border-slate-200 dark:border-slate-800">
                                        <div className="flex justify-between items-center p-5 border-b border-slate-100 dark:border-slate-700/50 bg-slate-50/50 dark:bg-slate-900/50">
                                            <h3 className="text-xl font-bold text-slate-800 dark:text-slate-100 flex items-center gap-2"><Icon name="settings" size={20} className="text-slate-400" /> Kelola Daftar Sertifikat</h3>
                                            <button onClick={() => setShowCertManager(false)} type="button" className="text-slate-400 hover:bg-slate-200 hover:text-slate-600 dark:hover:bg-slate-700 dark:hover:text-slate-300 p-2 rounded-xl transition-colors"><Icon name="x" size={20} /></button>
                                        </div>
                                        <div className="p-6 overflow-y-auto custom-scrollbar flex-1">
                                            <form onSubmit={(e) => {
                                                e.preventDefault();
                                                const input = e.target.elements.newcert.value.trim();
                                                if (!input) return;
                                                if (certList.includes(input)) {
                                                    setAlertModal({ isOpen: true, title: 'Duplikasi Sertifikat', message: "Sertifikat sudah ada di dalam daftar!" });
                                                    return;
                                                }
                                                handleUpdateCertList([input, ...certList]);
                                                e.target.reset();
                                            }} className="flex gap-2 mb-4">
                                                <input type="text" name="newcert" placeholder="Tambah nama sertifikat baru..." className="flex-1 p-2 border border-slate-300 dark:border-slate-700 rounded-lg outline-none focus:border-indigo-500 bg-slate-50 dark:bg-slate-800 text-sm dark:text-slate-200" />
                                                <button type="submit" className="px-3 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg text-sm font-semibold transition-colors flex items-center gap-1"><Icon name="plus" size={16} /> Tambah</button>
                                            </form>
                                            <div className="space-y-2">
                                                {certList.length === 0 ? <p className="text-sm text-slate-400 italic">Daftar kosong.</p> : null}
                                                {certList.map((cert, idx) => (
                                                    <div key={idx} className="flex justify-between items-center p-2.5 bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg">
                                                        <span className="text-sm text-slate-700 dark:text-slate-300">{cert}</span>
                                                        <button onClick={() => {
                                                            if (confirm(`Hapus ${cert} dari daftar sertifikat?`)) {
                                                                handleUpdateCertList(certList.filter(item => item !== cert));
                                                            }
                                                        }} className="text-slate-400 hover:text-red-500 transition-colors p-1" title="Hapus"><Icon name="trash-2" size={16} /></button>
                                                    </div>
                                                ))}
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            )}

                            {/* Alert Modal */}
                            {alertModal.isOpen && (
                                <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[200] flex items-center justify-center fade-in p-4">
                                    <div className="glass-card rounded-3xl shadow-2xl w-full max-w-sm p-6 transform scale-in border border-slate-200 dark:border-slate-800 text-center">
                                        <div className="w-16 h-16 bg-red-100 dark:bg-red-900/30 text-red-600 dark:text-red-400 rounded-full flex items-center justify-center mx-auto mb-4">
                                            <Icon name="alert-triangle" size={32} />
                                        </div>
                                        <h3 className="text-lg font-bold text-slate-800 dark:text-slate-100 mb-2">{alertModal.title}</h3>
                                        <p className="text-sm text-slate-600 dark:text-slate-400 mb-6 whitespace-pre-line">{alertModal.message}</p>
                                        <button onClick={() => setAlertModal({ isOpen: false, title: '', message: '' })} className="w-full px-4 py-3 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl font-bold transition-all shadow-md hover:shadow-lg hover:shadow-indigo-500/30 hover:-translate-y-0.5 active:translate-y-0">
                                            Mengerti
                                        </button>
                                    </div>
                                </div>
                            )}

                            {/* Floating Contextual Mobile Menu Overlay */}
                            {mobileMenuOpen && (
                                <div className="lg:hidden fixed inset-0 z-[80] flex justify-center items-end pb-28 animate-in fade-in duration-300">
                                    <div className="absolute inset-0 bg-slate-900/20 backdrop-blur-sm transition-opacity" onClick={() => setMobileMenuOpen(false)}></div>
                                    
                                    <div className="relative w-[90%] max-w-[380px] bg-white/80 dark:bg-slate-900/80 backdrop-blur-2xl border border-white/50 dark:border-slate-700/50 rounded-[32px] p-5 shadow-[0_20px_40px_-10px_rgba(0,0,0,0.2)] dark:shadow-indigo-900/20 animate-in slide-in-from-bottom-10 zoom-in-95 duration-300">
                                        <div className="flex justify-between items-center mb-4 px-1">
                                            <h3 className="text-lg font-black text-slate-800 dark:text-slate-100 tracking-tight">Menu Lainnya</h3>
                                            <button onClick={() => setMobileMenuOpen(false)} className="w-8 h-8 flex items-center justify-center rounded-full bg-slate-100/50 dark:bg-slate-800/50 text-slate-500 hover:text-slate-700 hover:bg-slate-200 transition-colors">
                                                <Icon name="x" size={18} />
                                            </button>
                                        </div>
                                        <div className="grid grid-cols-3 gap-2">
                                            <MobileMenuItem icon={<Icon name="calendar" size={20} />} label="Master Schedule" isActive={activeTab === 'master-schedule'} onClick={() => { handleTabChange('master-schedule'); setMobileMenuOpen(false); }} />
                                            <MobileMenuItem icon={<Icon name="award" size={20} />} label="Tenaga Ahli" isActive={activeTab === 'ahli'} onClick={() => { handleTabChange('ahli'); setMobileMenuOpen(false); }} />
                                            <MobileMenuItem icon={<Icon name="briefcase" size={20} />} label="Tugas Ahli" isActive={activeTab === 'penugasan'} onClick={() => { handleTabChange('penugasan'); setMobileMenuOpen(false); }} />
                                            <MobileMenuItem icon={<Icon name="calendar-days" size={20} />} label="Ploting" isActive={activeTab === 'gantt'} onClick={() => { handleTabChange('gantt'); setMobileMenuOpen(false); }} />
                                            <MobileMenuItem icon={<Icon name="box" size={20} />} label="Inventaris" isActive={activeTab === 'inventaris'} onClick={() => { handleTabChange('inventaris'); setMobileMenuOpen(false); }} />
                                            <MobileMenuItem icon={<Icon name="bar-chart" size={20} />} label="KPI" isActive={activeTab === 'kpi'} onClick={() => { handleTabChange('kpi'); setMobileMenuOpen(false); }} />
                                        </div>
                                    </div>
                                </div>
                            )}

                            <nav className="lg:hidden fixed bottom-6 left-1/2 -translate-x-1/2 h-[70px] w-[90%] max-w-[380px] bg-white/70 dark:bg-slate-900/70 backdrop-blur-3xl border border-white/50 dark:border-slate-700/50 rounded-[35px] z-[90] flex justify-between items-center px-3 shadow-[0_20px_40px_-10px_rgba(0,0,0,0.15)] dark:shadow-indigo-900/20 transition-all duration-500">
                                <BottomNavItem icon={<Icon name="layout-dashboard" size={20} />} label="Beranda" isActive={activeTab === 'dashboard' && !mobileMenuOpen} onClick={() => { handleTabChange('dashboard'); setMobileMenuOpen(false); }} />
                                <BottomNavItem icon={<Icon name="briefcase" size={20} />} label="Proyek" isActive={activeTab === 'proyek' && !mobileMenuOpen} onClick={() => { handleTabChange('proyek'); setMobileMenuOpen(false); }} />
                                <BottomNavItem icon={<Icon name="users" size={20} />} label="Tim" isActive={activeTab === 'tim' && !mobileMenuOpen} onClick={() => { handleTabChange('tim'); setMobileMenuOpen(false); }} />
                                <BottomNavItem icon={<Icon name="menu" size={20} />} label="Menu" isActive={mobileMenuOpen} onClick={() => setMobileMenuOpen(!mobileMenuOpen)} />
                            </nav>
                        </div>
                    </div>
                </>
            );
        }

        function MobileMenuItem({ icon, label, isActive, onClick }) {
            return (
                <button onClick={onClick} className={`flex flex-col items-center justify-center w-full p-3 gap-2 transition-all rounded-[20px] ${isActive ? 'bg-indigo-50 dark:bg-indigo-500/20 text-indigo-700 dark:text-indigo-300 shadow-sm' : 'hover:bg-slate-50 dark:hover:bg-slate-800/50 text-slate-600 dark:text-slate-300 hover:text-indigo-600 dark:hover:text-indigo-400'}`}>
                    <div className={`w-12 h-12 rounded-2xl flex items-center justify-center transition-transform duration-300 ${isActive ? 'bg-indigo-600 text-white scale-110 shadow-md shadow-indigo-600/30' : 'bg-white dark:bg-slate-800 shadow-sm border border-slate-100 dark:border-slate-700'}`}>
                        {icon}
                    </div>
                    <span className="text-[10px] font-bold text-center leading-tight">{label}</span>
                </button>
            );
        }

        function SidebarItem({ icon, label, isActive, onClick }) {
            return (
                <button onClick={onClick} className={`w-full flex items-center gap-3 px-4 py-3 rounded-2xl transition-all duration-300 font-semibold ${isActive ? 'bg-slate-800 dark:bg-white text-white dark:text-slate-900 shadow-lg shadow-slate-900/20 dark:shadow-white/20 scale-[1.02]' : 'text-slate-500 dark:text-slate-400 hover:bg-white/60 dark:hover:bg-slate-800 hover:text-slate-900 dark:hover:text-white'}`}>
                    {icon} <span className="font-medium text-sm tracking-wide">{label}</span>
                </button>
            );
        }

        function BottomNavItem({ icon, label, isActive, onClick }) {
            return (
                <button onClick={onClick} className={`relative flex items-center justify-center transition-all duration-500 ease-out h-[48px] overflow-hidden ${isActive ? 'bg-slate-800 dark:bg-white text-white dark:text-slate-900 rounded-full px-5 gap-2 shadow-lg shadow-slate-900/20 dark:shadow-white/20 w-auto' : 'w-12 text-slate-500 hover:text-slate-900 dark:text-slate-400 dark:hover:text-white bg-transparent rounded-full'}`}>
                    <div className={`transition-transform duration-500 ${isActive ? 'scale-100' : 'scale-95'}`}>{icon}</div>
                    {isActive && <span className="text-[12px] font-bold whitespace-nowrap animate-in fade-in zoom-in duration-300 tracking-wide">{label}</span>}
                </button>
            );
        }

        function MetricCard({ title, value, trend, trendUp, icon, color, onClick }) {
            const darkColorMap = {
                'bg-blue-50': 'dark:bg-blue-900/20 dark:text-blue-400',
                'bg-indigo-50': 'dark:bg-indigo-900/20 dark:text-indigo-400',
                'bg-amber-50': 'dark:bg-amber-900/20 dark:text-amber-400',
                'bg-emerald-50': 'dark:bg-emerald-900/20 dark:text-emerald-400',
                'bg-teal-50': 'dark:bg-teal-900/20 dark:text-teal-400',
            };
            const darkColor = darkColorMap[color] || '';

            return (
                <div onClick={onClick} className={`glass-card p-4 sm:p-5 lg:p-6 rounded-[2rem] hover:shadow-2xl transition-all duration-300 hover:-translate-y-1.5 flex flex-col justify-between ${onClick ? 'cursor-pointer' : ''}`}>
                    <div className="flex flex-wrap justify-between items-start gap-2 mb-4">
                        <div className={`p-3.5 rounded-2xl ${color} ${darkColor} shadow-inner`}>{icon}</div>
                        {trendUp !== null && (
                            <div className={`flex items-center gap-1 text-[10px] sm:text-xs font-bold px-2.5 py-1 rounded-full whitespace-nowrap ${trendUp ? 'bg-emerald-500/10 text-emerald-600 dark:bg-emerald-500/20 dark:text-emerald-400' : 'bg-red-500/10 text-red-600 dark:bg-red-500/20 dark:text-red-400'}`}>
                                <Icon name={trendUp ? 'trending-up' : 'trending-down'} size={12} /> {trend}
                            </div>
                        )}
                    </div>
                    <div>
                        <h4 className="text-slate-500 dark:text-slate-400 text-sm font-semibold mb-1 tracking-wide">{title}</h4>
                        <p className="text-2xl lg:text-3xl font-black text-slate-800 dark:text-white tracking-tight">{value}</p>
                    </div>
                </div>
            );
        }

        function StatusBadge({ status }) {
            let style = 'bg-slate-500/10 text-slate-700 dark:text-slate-300';
            if (status === 'Done') style = 'bg-emerald-500/10 text-emerald-700 dark:bg-emerald-500/20 dark:text-emerald-400';
            else if (status === 'On Progress') style = 'bg-blue-500/10 text-blue-700 dark:bg-blue-500/20 dark:text-blue-400';
            else if (status === 'Terlambat') style = 'bg-red-500/10 text-red-700 dark:bg-red-500/20 dark:text-red-400';
            else if (status === 'Beresiko') style = 'bg-amber-500/10 text-amber-700 dark:bg-amber-500/20 dark:text-amber-400';
            else if (status === 'Not Started') style = 'bg-slate-500/10 text-slate-600 dark:bg-slate-500/20 dark:text-slate-400';

            return (
                <span className={`px-2.5 py-1 rounded-full text-[11px] font-bold tracking-wide ${style} flex items-center gap-1.5 w-max`}>
                    {status === 'Done' && <Icon name="check-circle-2" size={12} />}
                    {status === 'On Progress' && <Icon name="clock" size={12} />}
                    {status === 'Terlambat' && <Icon name="alert-triangle" size={12} />}
                    {status === 'Beresiko' && <Icon name="clock" size={12} />}
                    {status === 'Not Started' && <Icon name="calendar" size={12} />}
                    {status}
                </span>
            );
        }

        
        
export default App;
