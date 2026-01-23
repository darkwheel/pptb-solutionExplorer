/// <reference types="@pptb/types" />

/**
 * Solution Explorer Tool for Power Platform Tool Box
 * 
 * This tool helps administrators explore and manage Dataverse solutions
 * with comprehensive metadata including publisher info, versions, deployment details, etc.
 */

// Global API references
const toolbox = window.toolboxAPI;
const dataverse = window.dataverseAPI;

// Application state
let currentConnection: ToolBoxAPI.DataverseConnection | null = null;
let allSolutions: any[] = [];
let filteredSolutions: any[] = [];
let selectedSolutionId: string | null = null;
let currentLimit: string = '100';

/**
 * Initialize the application
 */
async function initialize() {
    log('Solution Explorer initialized', 'info');
    
    try {
        // Check connection
        await refreshConnection();
        
        // Setup event handlers
        setupEventHandlers();
        
        // Subscribe to events
        subscribeToEvents();
        
        log('Tool initialized successfully', 'success');
    } catch (error) {
        log(`Initialization error: ${(error as Error).message}`, 'error');
        await toolbox.utils.showNotification({
            title: 'Initialization Error',
            body: (error as Error).message,
            type: 'error',
            duration: 3000
        });
    }
}

/**
 * Refresh connection information
 */
async function refreshConnection() {
    try {
        currentConnection = await toolbox.connections.getActiveConnection();
        
        // Update page title with connection name
        const titleElement = document.getElementById('page-title');
        if (titleElement) {
            if (currentConnection) {
                titleElement.textContent = `📦 Solution Explorer (${currentConnection.name})`;
                log(`Connected to: ${currentConnection.name}`, 'success');
            } else {
                titleElement.textContent = '📦 Solution Explorer';
                log('No active connection', 'warning');
            }
        }
    } catch (error) {
        log(`Error refreshing connection: ${(error as Error).message}`, 'error');
    }
}

/**
 * Subscribe to platform events
 */
function subscribeToEvents() {
    toolbox.events.on((event, payload) => {
        switch (payload.event) {
            case 'connection:updated':
            case 'connection:created':
                refreshConnection();
                break;
            case 'connection:deleted':
                currentConnection = null;
                refreshConnection();
                break;
        }
    });
}

/**
 * Setup UI event handlers
 */
function setupEventHandlers() {
    // Load solutions button
    document.getElementById('load-solutions-btn')?.addEventListener('click', loadSolutions);
    
    // Refresh button
    document.getElementById('refresh-btn')?.addEventListener('click', loadSolutions);
    
    // Clear log button
    document.getElementById('clear-log-btn')?.addEventListener('click', clearLog);
    
    // Search input
    const searchInput = document.getElementById('search-input') as HTMLInputElement;
    searchInput?.addEventListener('input', applyFilters);
    
    // Filter selects
    document.getElementById('type-filter')?.addEventListener('change', applyFilters);
    document.getElementById('visibility-filter')?.addEventListener('change', applyFilters);
    
    // Limit filter - triggers reload
    document.getElementById('limit-filter')?.addEventListener('change', (e) => {
        currentLimit = (e.target as HTMLSelectElement).value;
        loadSolutions();
    });
}

/**
 * Load all solutions from Dataverse
 */
async function loadSolutions() {
    if (!currentConnection) {
        await toolbox.utils.showNotification({
            title: 'No Connection',
            body: 'Please connect to a Dataverse environment',
            type: 'warning',
            duration: 3000
        });
        return;
    }
    
    try {
        const loadBtn = document.getElementById('load-solutions-btn') as HTMLButtonElement;
        if (loadBtn) loadBtn.disabled = true;
        
        log('Loading solutions...', 'info');
        
        // Show loading
        await toolbox.utils.showLoading('Loading solutions...');
        
        // FetchXML to get solutions with publisher information (dynamically set top value)
        const topAttribute = currentLimit === 'all' ? '' : `top="${currentLimit}"`;
        const fetchXml = `
<fetch ${topAttribute}>
    <entity name="solution">
        <attribute name="solutionid" />
        <attribute name="uniquename" />
        <attribute name="friendlyname" />
        <attribute name="version" />
        <attribute name="description" />
        <attribute name="installedon" />
        <attribute name="createdon" />
        <attribute name="modifiedon" />
        <attribute name="ismanaged" />
        <attribute name="isvisible" />
        <attribute name="solutionpackageversion" />
        <attribute name="solutiontype" />
        <attribute name="publisherid" />
        <link-entity name="publisher" from="publisherid" to="publisherid" alias="publisher">
            <attribute name="friendlyname" />
            <attribute name="uniquename" />
            <attribute name="customizationprefix" />
            <attribute name="customizationoptionvalueprefix" />
            <attribute name="description" />
        </link-entity>
        <link-entity name="systemuser" from="systemuserid" to="modifiedby" alias="modifiedby">
            <attribute name="fullname" />
            <attribute name="domainname" />
        </link-entity>
        <filter>
            <condition attribute="isvisible" operator="eq" value="1" />
        </filter>
        <order attribute="installedon" descending="true" />
    </entity>
</fetch>`;
        
        const result = await dataverse.fetchXmlQuery(fetchXml);
        allSolutions = result.value;
        
        log(`Loaded ${allSolutions.length} solution(s)`, 'success');
        
        // Apply filters
        applyFilters();
        
        // Show solutions section
        const solutionsSection = document.getElementById('solutions-section');
        if (solutionsSection) solutionsSection.style.display = 'block';
        
        await toolbox.utils.showNotification({
            title: 'Solutions Loaded',
            body: `Successfully loaded ${allSolutions.length} solution(s)`,
            type: 'success',
            duration: 3000
        });
        
    } catch (error) {
        log(`Error loading solutions: ${(error as Error).message}`, 'error');
        await toolbox.utils.showNotification({
            title: 'Load Failed',
            body: (error as Error).message,
            type: 'error',
            duration: 5000
        });
    } finally {
        await toolbox.utils.hideLoading();
        const loadBtn = document.getElementById('load-solutions-btn') as HTMLButtonElement;
        if (loadBtn) loadBtn.disabled = false;
    }
}

/**
 * Apply filters to solutions list
 */
function applyFilters() {
    const searchInput = document.getElementById('search-input') as HTMLInputElement;
    const typeFilter = document.getElementById('type-filter') as HTMLSelectElement;
    const visibilityFilter = document.getElementById('visibility-filter') as HTMLSelectElement;
    
    const searchTerm = searchInput?.value.toLowerCase() || '';
    const typeValue = typeFilter?.value || 'all';
    const visibilityValue = visibilityFilter?.value || 'all';
    
    filteredSolutions = allSolutions.filter(solution => {
        // Search filter
        const matchesSearch = !searchTerm || 
            (solution.friendlyname || '').toLowerCase().includes(searchTerm) ||
            (solution.uniquename || '').toLowerCase().includes(searchTerm) ||
            (solution['publisher.friendlyname'] || '').toLowerCase().includes(searchTerm);
        
        // Type filter
        const matchesType = typeValue === 'all' ||
            (typeValue === 'managed' && solution.ismanaged) ||
            (typeValue === 'unmanaged' && !solution.ismanaged);
        
        // Visibility filter
        const matchesVisibility = visibilityValue === 'all' ||
            (visibilityValue === 'visible' && solution.isvisible) ||
            (visibilityValue === 'hidden' && !solution.isvisible);
        
        return matchesSearch && matchesType && matchesVisibility;
    });
    
    displaySolutions();
}

/**
 * Display solutions list
 */
function displaySolutions() {
    const listDiv = document.getElementById('solutions-list');
    const countDiv = document.getElementById('solutions-count');
    
    if (!listDiv) return;
    
    if (countDiv) {
        countDiv.textContent = `${filteredSolutions.length} solution(s) ${filteredSolutions.length !== allSolutions.length ? `(filtered from ${allSolutions.length})` : ''}`;
    }
    
    if (filteredSolutions.length === 0) {
        listDiv.innerHTML = '<div class="empty-message">No solutions found matching your criteria</div>';
        return;
    }
    
    listDiv.innerHTML = filteredSolutions.map(solution => {
        const isManaged = solution.ismanaged;
        const typeLabel = isManaged ? 'Managed' : 'Unmanaged';
        const typeClass = isManaged ? 'managed' : 'unmanaged';
        
        return `
            <div class="solution-card ${typeClass}" data-solutionid="${solution.solutionid}">
                <div class="solution-header">
                    <div class="solution-name">${escapeHtml(solution.friendlyname || solution.uniquename || 'N/A')}</div>
                    <span class="solution-type-badge ${typeClass}">${typeLabel}</span>
                </div>
                <div class="solution-meta">
                    <div class="solution-version">v${escapeHtml(solution.version || '1.0.0.0')}</div>
                    <div class="solution-publisher">${escapeHtml(solution['publisher.friendlyname'] || 'N/A')}</div>
                </div>
                <div class="solution-unique-name">${escapeHtml(solution.uniquename || 'N/A')}</div>
            </div>
        `;
    }).join('');
    
    // Add click handlers
    document.querySelectorAll('.solution-card').forEach(card => {
        card.addEventListener('click', () => {
            const solutionId = (card as HTMLElement).getAttribute('data-solutionid');
            if (solutionId) {
                const solution = filteredSolutions.find(s => s.solutionid === solutionId);
                if (solution) {
                    selectSolution(solutionId, solution);
                }
            }
        });
    });
}

/**
 * Select a solution and display details
 */
async function selectSolution(solutionId: string, solutionData: any) {
    selectedSolutionId = solutionId;
    
    // Update selected state
    document.querySelectorAll('.solution-card').forEach(card => {
        card.classList.remove('selected');
        if ((card as HTMLElement).getAttribute('data-solutionid') === solutionId) {
            card.classList.add('selected');
        }
    });
    
    const detailsPanel = document.getElementById('solution-details-panel');
    if (!detailsPanel) return;
    
    // Show loading state
    detailsPanel.innerHTML = '<div class="loading-indicator">Loading solution details...</div>';
    
    try {
        // Get component count
        const componentCount = await getSolutionComponentCount(solutionId);
        
        // Get deployment history
        const deploymentHistory = await getSolutionDeploymentHistory(solutionId);
        
        // Display details
        displaySolutionDetails(solutionData, componentCount, deploymentHistory);
        
        log(`Selected solution: ${solutionData.friendlyname}`, 'info');
        
    } catch (error) {
        log(`Error loading solution details: ${(error as Error).message}`, 'error');
        detailsPanel.innerHTML = `<div class="empty-state"><p>Error loading details: ${(error as Error).message}</p></div>`;
    }
}

/**
 * Get solution component count
 */
async function getSolutionComponentCount(solutionId: string): Promise<number> {
    try {
        const fetchXml = `
<fetch aggregate="true">
    <entity name="solutioncomponent">
        <attribute name="solutioncomponentid" alias="count" aggregate="count" />
        <filter>
            <condition attribute="solutionid" operator="eq" value="${solutionId}" />
        </filter>
    </entity>
</fetch>`;
        
        const result = await dataverse.fetchXmlQuery(fetchXml);
        return parseInt((result.value[0]?.count as string) || '0', 10);
    } catch (error) {
        log(`Error getting component count: ${(error as Error).message}`, 'warning');
        return 0;
    }
}

/**
 * Get deployment history for a solution
 */
async function getSolutionDeploymentHistory(solutionId: string): Promise<any[]> {
    try {
        const fetchXml = `
<fetch top="20">
    <entity name="importjob">
        <attribute name="importjobid" />
        <attribute name="solutionname" />
        <attribute name="completedon" />
        <attribute name="startedon" />
        <attribute name="progress" />
        <attribute name="createdonbehalfby" />
        <attribute name="createdby" />
        <attribute name="modifiedby" />
        <filter>
            <condition attribute="solutionid" operator="eq" value="${solutionId}" />
        </filter>
        <order attribute="completedon" descending="true" />
    </entity>
</fetch>`;
        
        const result = await dataverse.fetchXmlQuery(fetchXml);
        return result.value || [];
    } catch (error) {
        log(`Error getting deployment history: ${(error as Error).message}`, 'warning');
        return [];
    }
}

/**
 * Display solution details in the right panel
 */
function displaySolutionDetails(solution: any, componentCount: number, deploymentHistory: any[] = []) {
    const detailsPanel = document.getElementById('solution-details-panel');
    if (!detailsPanel) return;
    
    const isManaged = solution.ismanaged;
    const typeLabel = isManaged ? 'Managed' : 'Unmanaged';
    const typeClass = isManaged ? 'managed' : 'unmanaged';
    
    detailsPanel.innerHTML = `
        <div class="solution-details">
            <div class="details-header">
                <h3>${escapeHtml(solution.friendlyname || solution.uniquename || 'N/A')}</h3>
                <span class="solution-type-badge ${typeClass}">${typeLabel}</span>
            </div>
            
            <div class="details-section">
                <h4>Basic Information</h4>
                <table class="details-table">
                    <tr>
                        <td class="label">Display Name:</td>
                        <td>${escapeHtml(solution.friendlyname || 'N/A')}</td>
                    </tr>
                    <tr>
                        <td class="label">Unique Name:</td>
                        <td><code>${escapeHtml(solution.uniquename || 'N/A')}</code></td>
                    </tr>
                    <tr>
                        <td class="label">Version:</td>
                        <td>${escapeHtml(solution.version || '1.0.0.0')}</td>
                    </tr>
                    <tr>
                        <td class="label">Description:</td>
                        <td>${escapeHtml(solution.description || 'No description provided')}</td>
                    </tr>
                    <tr>
                        <td class="label">Solution ID:</td>
                        <td><code>${escapeHtml(solution.solutionid || 'N/A')}</code></td>
                    </tr>
                </table>
            </div>
            
            <div class="details-section">
                <h4>Publisher Information</h4>
                <table class="details-table">
                    <tr>
                        <td class="label">Publisher Name:</td>
                        <td>${escapeHtml(solution['publisher.friendlyname'] || 'N/A')}</td>
                    </tr>
                    <tr>
                        <td class="label">Publisher Unique Name:</td>
                        <td><code>${escapeHtml(solution['publisher.uniquename'] || 'N/A')}</code></td>
                    </tr>
                    <tr>
                        <td class="label">Customization Prefix:</td>
                        <td><code>${escapeHtml(solution['publisher.customizationprefix'] || 'N/A')}</code></td>
                    </tr>
                    <tr>
                        <td class="label">Option Value Prefix:</td>
                        <td>${solution['publisher.customizationoptionvalueprefix'] !== null && solution['publisher.customizationoptionvalueprefix'] !== undefined ? solution['publisher.customizationoptionvalueprefix'] : 'N/A'}</td>
                    </tr>
                    ${solution['publisher.description'] ? `
                    <tr>
                        <td class="label">Publisher Description:</td>
                        <td>${escapeHtml(solution['publisher.description'])}</td>
                    </tr>
                    ` : ''}
                </table>
            </div>
            
            <div class="details-section">
                <h4>Installation & Deployment</h4>
                <table class="details-table">
                    <tr>
                        <td class="label">Installed On:</td>
                        <td>${solution.installedon ? formatDateTime(solution.installedon) : 'N/A'}</td>
                    </tr>
                    <tr>
                        <td class="label">Created On:</td>
                        <td>${solution.createdon ? formatDateTime(solution.createdon) : 'N/A'}</td>
                    </tr>
                    <tr>
                        <td class="label">Modified On:</td>
                        <td>${solution.modifiedon ? formatDateTime(solution.modifiedon) : 'N/A'}</td>
                    </tr>
                    <tr>
                        <td class="label">Last Modified By:</td>
                        <td>
                            <div class="modifier-info">
                                <div class="modifier-name">${escapeHtml(solution['modifiedby.fullname'] || 'N/A')}</div>
                                ${solution['modifiedby.domainname'] ? `<div class="modifier-domain">${escapeHtml(solution['modifiedby.domainname'])}</div>` : ''}
                            </div>
                        </td>
                    </tr>
                    <tr>
                        <td class="label">Package Type:</td>
                        <td>${isManaged ? 'Managed' : 'Unmanaged'}</td>
                    </tr>
                    <tr>
                        <td class="label">Is Visible:</td>
                        <td>${solution.isvisible ? 'Yes' : 'No'}</td>
                    </tr>
                    ${solution.solutionpackageversion ? `
                    <tr>
                        <td class="label">Package Version:</td>
                        <td>${escapeHtml(solution.solutionpackageversion)}</td>
                    </tr>
                    ` : ''}
                </table>
            </div>
            
            <div class="details-section">
                <h4>Components</h4>
                <table class="details-table">
                    <tr>
                        <td class="label">Component Count:</td>
                        <td><strong>${componentCount}</strong> component(s)</td>
                    </tr>
                </table>
            </div>
            
            ${deploymentHistory.length > 0 ? `
            <div class="details-section">
                <h4>Deployment History</h4>
                <div class="deployment-history">
                    ${deploymentHistory.map((deployment: any, index: number) => {
                        const completedOn = deployment.completedon ? formatDateTime(deployment.completedon) : 'N/A';
                        const startedOn = deployment.startedon ? formatDateTime(deployment.startedon) : 'N/A';
                        const deployedBy = escapeHtml(deployment['_createdby_value@OData.Community.Display.V1.FormattedValue'] || 'System');
                        const progress = deployment.progress !== null && deployment.progress !== undefined ? Math.round(deployment.progress) : 100;
                        
                        return `
                        <div class="deployment-item">
                            <div class="deployment-header">
                                <div class="deployment-number">#${index + 1}</div>
                                <div class="deployment-date">${completedOn}</div>
                            </div>
                            <div class="deployment-details">
                                <div class="deployment-info">
                                    <span class="label">Deployed By:</span>
                                    <span class="value">${deployedBy}</span>
                                </div>
                                <div class="deployment-info">
                                    <span class="label">Started:</span>
                                    <span class="value">${startedOn}</span>
                                </div>
                                <div class="deployment-info">
                                    <span class="label">Progress:</span>
                                    <span class="value">${progress}%</span>
                                </div>
                            </div>
                        </div>
                        `;
                    }).join('')}
                </div>
            </div>
            ` : ''}
            
            <div class="details-section">
                <h4>Management Actions</h4>
                <div class="action-buttons">
                    <button class="btn btn-secondary btn-small" onclick="copySolutionId('${solution.solutionid}')">Copy Solution ID</button>
                    <button class="btn btn-secondary btn-small" onclick="copySolutionUniqueName('${solution.uniquename}')">Copy Unique Name</button>
                </div>
            </div>
        </div>
    `;
}

/**
 * Copy solution ID to clipboard
 */
async function copySolutionId(solutionId: string) {
    try {
        await toolbox.utils.copyToClipboard(solutionId);
        await toolbox.utils.showNotification({
            title: 'Copied',
            body: 'Solution ID copied to clipboard',
            type: 'success',
            duration: 2000
        });
        log('Solution ID copied to clipboard', 'info');
    } catch (error) {
        log(`Error copying to clipboard: ${(error as Error).message}`, 'error');
    }
}

/**
 * Copy solution unique name to clipboard
 */
async function copySolutionUniqueName(uniqueName: string) {
    try {
        await toolbox.utils.copyToClipboard(uniqueName);
        await toolbox.utils.showNotification({
            title: 'Copied',
            body: 'Unique name copied to clipboard',
            type: 'success',
            duration: 2000
        });
        log('Unique name copied to clipboard', 'info');
    } catch (error) {
        log(`Error copying to clipboard: ${(error as Error).message}`, 'error');
    }
}

// Make functions globally available for onclick handlers
(window as any).copySolutionId = copySolutionId;
(window as any).copySolutionUniqueName = copySolutionUniqueName;

/**
 * Format date/time for display
 */
function formatDateTime(dateString: string): string {
    try {
        const date = new Date(dateString);
        return date.toLocaleString();
    } catch {
        return dateString;
    }
}

/**
 * Escape HTML to prevent XSS
 */
function escapeHtml(text: string): string {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

/**
 * Log message to event log
 */
function log(message: string, type: 'info' | 'success' | 'warning' | 'error' = 'info') {
    const logDiv = document.getElementById('event-log');
    if (!logDiv) return;
    
    const timestamp = new Date().toLocaleTimeString();
    const logEntry = document.createElement('div');
    logEntry.className = `log-entry ${type}`;
    logEntry.innerHTML = `
        <span class="log-timestamp">[${timestamp}]</span>
        <span>${message}</span>
    `;
    
    logDiv.insertBefore(logEntry, logDiv.firstChild);
    
    // Keep only last 50 entries
    while (logDiv.children.length > 50) {
        logDiv.removeChild(logDiv.lastChild!);
    }
    
    console.log(`[${type.toUpperCase()}] ${message}`);
}

/**
 * Clear event log
 */
function clearLog() {
    const logDiv = document.getElementById('event-log');
    if (logDiv) {
        logDiv.innerHTML = '';
        log('Log cleared', 'info');
    }
}

// Initialize when DOM is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initialize);
} else {
    initialize();
}
