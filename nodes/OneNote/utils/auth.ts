/** Shared helpers for the "device login" authentication (no redirect URI needed). */

export type Tenant = 'consumers' | 'common';

export function authority(tenant: string): string {
	return `https://login.microsoftonline.com/${tenant || 'consumers'}/oauth2/v2.0`;
}

export function scopes(enableOneDrive: boolean): string {
	return `offline_access User.Read Notes.ReadWrite${enableOneDrive ? ' Files.ReadWrite' : ''}`;
}

export function formBody(values: Record<string, string>): string {
	return new URLSearchParams(values).toString();
}

export const FORM_HEADERS = { 'Content-Type': 'application/x-www-form-urlencoded' };
