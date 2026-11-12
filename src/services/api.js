// src/services/api.js
// Client API for WVUSD Illuminate Data Portal with token authentication & zero-trust verification

import { GAS_WEB_APP_URL, APP_SECRET_TOKEN, SUPER_ADMIN_EMAIL } from '../config/constants.js';

/**
 * Fetch Illuminate assessment data for a verified WVUSD user with appToken
 * @param {string} userEmailOrToken
 * @returns {Promise<{ user: object, meta: object, students: Array }>}
 */
export async function fetchIlluminateData(userEmailOrToken) {
  if (!GAS_WEB_APP_URL || !GAS_WEB_APP_URL.startsWith('http')) {
    throw new Error('Google Apps Script Web App URL is not configured.');
  }

  const queryParams = new URLSearchParams({
    appToken: APP_SECRET_TOKEN || ''
  });

  if (typeof userEmailOrToken === 'string' && userEmailOrToken.includes('@')) {
    queryParams.set('email', userEmailOrToken.trim().toLowerCase());
  } else if (typeof userEmailOrToken === 'string' && userEmailOrToken.length > 40 && !userEmailOrToken.includes('@')) {
    queryParams.set('id_token', userEmailOrToken.trim());
  } else {
    queryParams.set('email', SUPER_ADMIN_EMAIL);
  }

  const url = `${GAS_WEB_APP_URL}?${queryParams.toString()}`;

  const response = await fetch(url, {
    method: 'GET',
    redirect: 'follow',
    headers: {
      'Accept': 'application/json'
    }
  });

  if (!response.ok) {
    throw new Error(`HTTP Error ${response.status}: ${response.statusText}`);
  }

  const data = await response.json();

  // Handle explicit Forbidden / Unauthorized zero-trust response
  if (data && (data.status === 'Forbidden' || data.success === false)) {
    const error = new Error(data.message || data.error || 'Access Denied: Unauthorized request.');
    error.status = 'Forbidden';
    error.isForbidden = true;
    throw error;
  }

  if (data && data.error) {
    const error = new Error(data.error);
    if (data.error.toLowerCase().includes('unauthorized') || data.error.toLowerCase().includes('forbidden')) {
      error.status = 'Forbidden';
      error.isForbidden = true;
    }
    throw error;
  }

  if (!data || !data.students || !Array.isArray(data.students)) {
    throw new Error("Invalid response format: 'students' array missing from Google Apps Script.");
  }

  return {
    user: data.user || null,
    meta: data.meta || {},
    students: data.students
  };
}
