/*
 * Copyright 2025 Adobe. All rights reserved.
 * This file is licensed to you under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License. You may obtain a copy
 * of the License at http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software distributed under
 * the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR REPRESENTATIONS
 * OF ANY KIND, either express or implied. See the License for the specific language
 * governing permissions and limitations under the License.
 */

import core from '@actions/core';
import path from 'path';

import { OPERATIONS } from './sta-aem-helper-constants.js';

/**
 * The publish or preview endpoint prefix for the HLX Admin API.
 */
export const HELIX_API_PREFIX = Object.freeze({
  PREVIEW: 'preview',
  LIVE: 'live',
});

/**
 * The Helix Admin API endpoint.
 */
const HELIX_ENDPOINT = 'https://admin.hlx.page';

/**
 * Helper function to take the input operation string
 * and turn it into a friendly string.
 * @returns
 */
const getOperationName = (operation) => {
  switch (operation) {
    case OPERATIONS.PREVIEW_PAGES:
      return 'preview';
    case OPERATIONS.PREVIEW_AND_PUBLISH:
      return 'preview and/or publish';
    default:
      return 'unknown';
  }
};

/**
 * Update the file path to satisfy the HLX Admin API requirements.
 * @param {string} filePath - The path to process.
 * @param {boolean} force - Whether to force extension removal.
 * @returns {string} The processed path.
 */
const fixPathForHelix = (filePath, force = false) => {
  const dir = path.dirname(filePath);
  const base = path.basename(filePath, path.extname(filePath));

  // HTML for DA
  // DOCS for Sharepoint
  core.info(`filePath to fix: ${filePath}`);

  if (force
    || filePath.endsWith('.docx')
    || filePath.endsWith('.html')) {
    return path.join(dir, base);
  } else if (filePath.endsWith('.xlsx')) {
    return path.join(dir, `${base}.json`);
  }
  return filePath;
};

/**
 * Performs the preview or publish operation.
 *
 * @param {string} apiEndpoint - The API endpoint to call.
 * @param {string} pagePath - The page path to preview or publish. /some/page.docx
 * @param {string} token - The DA access token.
 * @returns {Promise<boolean>} - Returns true if successful, false otherwise.
 */
async function performPreviewPublish(apiEndpoint, pagePath, token) {
  const action = new URL(apiEndpoint)
    .pathname
    .startsWith('/preview/')
    ? HELIX_API_PREFIX.PREVIEW
    : HELIX_API_PREFIX.LIVE;

  const page = fixPathForHelix(pagePath);

  core.info(`Prev/Pub: to ${apiEndpoint} with page ${page}`);

  try {
    const resp = await fetch(`${apiEndpoint}${page}`, {
      method: 'POST',
      body: '{}',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
        'Access-Control-Expose-Headers': 'x-error',
      },
    });

    if (!resp.ok) {
      const xError = resp.headers.get('x-error');
      core.info(`.${action} operation failed on ${page}: ${resp.status} : ${resp.statusText} : ${xError}`);

      // Check for unsupported media type or 404, and try without an extension
      if (resp.status === 415 || (action === 'live' && resp.status === 404)) {
        const noExtPath = fixPathForHelix(pagePath, true);
        // Avoid infinite loop by ensuring the path changed.
        if (noExtPath !== page && noExtPath !== pagePath) {
          core.info(`❓ Failed with an "Unsupported Media" or 404 error. Retrying operation without an extension: ${noExtPath}`);
          return performPreviewPublish(apiEndpoint, noExtPath);
        }
        core.warning(`❌ Operation failed on extensionless ${page}: ${xError}`);
      } else if (resp.status === 423) {
        core.warning(`❌ Operation failed on ${page}. The file appears locked. Is it being edited? (${xError})`);
      } else {
        core.warning(`❌ Operation failed on ${page}: ${xError}`);
      }
      return false;
    }

    const data = await resp.json();
    core.info(`✓ Operation successful on ${page}: ${data[action].url}`);
    return true;
  } catch (error) {
    core.warning(`❌ Operation call failed on ${page}: ${error.message}`);
  }

  return false;
}

/**
 * Performs the preview or publish pages operation and sets the outputs.
 * @param {string} pages - The URLs to preview or publish.
 * @param {string} operation - The operation to perform.
 * @param {string} context - The AEMY context.
 * @param {string} token - The DA access token.
 * @throws {Error} - If the operation fails.
 */
export async function doPreviewPublish(pages, operation, context, token) {
  const { project } = JSON.parse(context);
  const { owner, repo, branch = 'main' } = project;

  if (!owner || !repo) {
    throw new Error('Invalid context format: missing owner or repo.');
  }

  // keep track of the number of successes and failures
  const report = {
    successes: 0,
    failures: 0,
    failureList: {
      preview: [],
      publish: [],
    },
  };

  // if operation is OPERATIONS.PREVIEW_AND_PUBLISH we need to process the pages twice
  const loops = operation === OPERATIONS.PREVIEW_AND_PUBLISH ? 2 : 1;
  for (let i = 0; i < loops; i++) {
    // always preview first, then publish
    const action = [i === 0 ? HELIX_API_PREFIX.PREVIEW : HELIX_API_PREFIX.LIVE];
    const apiEndpoint = `${HELIX_ENDPOINT}/${action}/${owner}/${repo}/${branch}`;

    for (const page of pages) {
      const result = await performPreviewPublish(apiEndpoint, page, token);
      if (result) {
        report.successes += 1;
        core.info(`Increased successes to: ${report.successes}`);
      } else {
        report.failures += 1;
        report.failureList[action].push(page);
      }
    }
  }

  core.setOutput('successes', report.successes);
  core.setOutput('failures', report.failures);

  if (report.failures > 0) {
    core.warning(`❌ The pages that failed are: ${JSON.stringify(report.failureList, undefined, 2)}`);
    core.setOutput('error_message', `❌ Error: Failed to ${getOperationName(operation)}]} ${report.failures} of ${pages.length} pages.`);
    // eslint-disable-next-line max-len
  } else if (((operation === OPERATIONS.PREVIEW_AND_PUBLISH ? 2 : 1) * pages.length) !== report.successes) {
    core.warning(`❌ The paths that failed are: ${JSON.stringify(report.failureList, undefined, 2)}`);
    core.setOutput('error_message', `❌ Error: Failed to ${getOperationName(operation)} all of the paths.`);
  }
}
