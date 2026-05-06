/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export interface Student {
  id: string;
  name: string;
  class: string;
  contactNumber: string;
  shift?: '1st' | '2nd' | '3rd';
  createdAt: number;
}

export interface DailyReport {
  id: string;
  studentId: string;
  studentName?: string;
  date: string;
  attendance: boolean;
  homeworkStatus?: 'Complete' | 'Incomplete' | 'Partially Complete';
  classFocus?: 'Good' | 'Moderate' | 'Poor' | 'Extremely Poor';
  classBehaviour?: 'Excellent' | 'Good' | 'Normal' | 'Poor';
  complaint?: string;
  createdAt: number;
  createdBy: string;
}

export type TabType = 'home' | 'students' | 'reports';
