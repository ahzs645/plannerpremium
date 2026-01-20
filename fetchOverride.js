/**
 * Fetch Override for Planner Exporter
 * Intercepts fetch calls to capture API tokens from request headers
 * Supports both Graph API and PSS (Project Scheduling Service) API
 * This is critical because MSAL now encrypts tokens in localStorage
 */

(function() {
  'use strict';

  // Store captured tokens
  window.__plannerGraphToken = null;
  window.__plannerGraphTokenCapturedAt = null;

  // PSS API tokens (for Premium Plans)
  window.__plannerPssToken = null;
  window.__plannerPssTokenCapturedAt = null;
  window.__plannerPssProjectId = null;

  const originalFetch = window.fetch;

  // Helper to extract auth header from various formats
  function extractAuthHeader(config, resource) {
    let authHeader = null;

    if (config && config.headers) {
      if (config.headers instanceof Headers) {
        authHeader = config.headers.get('Authorization') || config.headers.get('authorization');
      } else if (Array.isArray(config.headers)) {
        const authEntry = config.headers.find(([key]) =>
          key.toLowerCase() === 'authorization'
        );
        if (authEntry) authHeader = authEntry[1];
      } else if (typeof config.headers === 'object') {
        authHeader = config.headers['Authorization'] || config.headers['authorization'];
      }
    }

    // Also check Request object headers
    if (!authHeader && resource instanceof Request) {
      authHeader = resource.headers.get('Authorization') || resource.headers.get('authorization');
    }

    return authHeader;
  }

  // Extract project ID from PSS API URL
  // Format: /projects('{projectId}')/tasks or /projects({projectId})/tasks
  // Also try to find msxrm_ pattern anywhere in URL
  function extractPssProjectId(url) {
    // Try standard /projects('id') pattern
    const match = url.match(/\/projects\(['"]?([^'")]+)['"]?\)/i);
    if (match) return match[1];

    // Try msxrm_ pattern anywhere in URL (query params, etc)
    const msxrmMatch = url.match(/msxrm_[^&"'\s]+/i);
    if (msxrmMatch) return msxrmMatch[0];

    return null;
  }

  // Extract Dynamics org from PSS API URL
  // Format: https://project.microsoft.com/orgXXX.crm3.dynamics.com/api/...
  function extractDynamicsOrg(url) {
    const match = url.match(/project\.microsoft\.com\/(org[a-f0-9]+\.crm\d*\.dynamics\.com)/i);
    return match ? match[1] : null;
  }

  window.fetch = async function(...args) {
    const [resource, config] = args;

    // Get URL string
    let url = '';
    if (typeof resource === 'string') {
      url = resource;
    } else if (resource instanceof Request) {
      url = resource.url;
    } else if (resource && resource.url) {
      url = resource.url;
    }

    // Check if this is a PSS API call (Premium Plans)
    if (url && url.includes('project.microsoft.com')) {
      const authHeader = extractAuthHeader(config, resource);

      if (authHeader && authHeader.startsWith('Bearer ')) {
        const token = authHeader.substring(7);

        if (token.split('.').length === 3) {
          window.__plannerPssToken = token;
          window.__plannerPssTokenCapturedAt = Date.now();

          // Extract project ID from URL
          const projectId = extractPssProjectId(url);
          if (projectId) {
            window.__plannerPssProjectId = projectId;
          }

          // Extract and store Dynamics org from URL
          const dynamicsOrg = extractDynamicsOrg(url);
          if (dynamicsOrg) {
            window.__plannerDynamicsOrg = dynamicsOrg;
            console.log('[Planner Exporter] Dynamics org captured:', dynamicsOrg);
          }

          // Notify via custom event
          window.dispatchEvent(new CustomEvent('planner-pss-token-captured', {
            detail: {
              token: token,
              projectId: projectId,
              timestamp: Date.now(),
              url: url
            }
          }));

          // Also post message for cross-context
          window.postMessage({
            type: 'PLANNER_PSS_TOKEN_CAPTURED',
            token: token,
            projectId: projectId,
            timestamp: Date.now()
          }, '*');

          console.log('[Planner Exporter] PSS token captured for project:', projectId, 'URL:', url);
        }
      }
    }

    // Check if this is a Graph API call or batch call
    if (url && (url.includes('graph.microsoft.com') || url.includes('$batch'))) {
      const authHeader = extractAuthHeader(config, resource);

      // Extract Bearer token
      if (authHeader && authHeader.startsWith('Bearer ')) {
        const token = authHeader.substring(7);

        // Validate it looks like a JWT (has 3 parts)
        if (token.split('.').length === 3) {
          window.__plannerGraphToken = token;
          window.__plannerGraphTokenCapturedAt = Date.now();

          // Notify via custom event
          window.dispatchEvent(new CustomEvent('planner-token-captured', {
            detail: {
              token: token,
              timestamp: Date.now(),
              url: url
            }
          }));

          // Also post message for cross-context
          window.postMessage({
            type: 'PLANNER_TOKEN_CAPTURED',
            token: token,
            timestamp: Date.now()
          }, '*');
        }
      }
    }

    // Call original fetch
    return originalFetch.apply(this, args);
  };

  // Also intercept XMLHttpRequest for completeness
  const originalXHROpen = XMLHttpRequest.prototype.open;
  const originalXHRSetHeader = XMLHttpRequest.prototype.setRequestHeader;

  XMLHttpRequest.prototype.open = function(method, url, ...rest) {
    this._plannerUrl = url;
    return originalXHROpen.apply(this, [method, url, ...rest]);
  };

  XMLHttpRequest.prototype.setRequestHeader = function(name, value) {
    if (this._plannerUrl && name.toLowerCase() === 'authorization' && value.startsWith('Bearer ')) {
      const token = value.substring(7);

      if (token.split('.').length === 3) {
        // PSS API (Premium Plans)
        if (this._plannerUrl.includes('project.microsoft.com')) {
          window.__plannerPssToken = token;
          window.__plannerPssTokenCapturedAt = Date.now();

          const projectId = extractPssProjectId(this._plannerUrl);
          if (projectId) {
            window.__plannerPssProjectId = projectId;
          }
        }

        // Graph API
        if (this._plannerUrl.includes('graph.microsoft.com') || this._plannerUrl.includes('$batch')) {
          window.__plannerGraphToken = token;
          window.__plannerGraphTokenCapturedAt = Date.now();
        }
      }
    }
    return originalXHRSetHeader.apply(this, [name, value]);
  };

  console.log('[Planner Exporter] Fetch override initialized (Graph + PSS)');
})();
