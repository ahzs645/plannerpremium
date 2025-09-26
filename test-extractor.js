/**
 * Test script to validate extraction patterns on actual Planner HTML files
 * This runs in Node.js environment using JSDOM to simulate browser DOM
 */

const fs = require('fs');
const path = require('path');

class PlannerTestExtractor {
  constructor() {
    this.results = {};
  }

  async testBothFiles() {
    console.log('🧪 Testing Planner Data Extraction Patterns\n');

    const cliFiles = process.argv.slice(2);
    const filesToTest = cliFiles.length > 0
      ? cliFiles.map((filename, index) => ({ filename, viewType: `file-${index + 1}` }))
      : [
          { filename: path.join('Tests', 'multi task plan.html'), viewType: 'generalView' }
        ];

    for (const { filename, viewType } of filesToTest) {
      await this.testFile(filename, viewType);
    }

    this.generateReport();
  }

  async testFile(filename, viewType) {
    console.log(`📄 Testing ${filename} (${viewType})`);
    console.log('=' .repeat(50));

    try {
      const htmlContent = fs.readFileSync(filename, 'utf8');
      const { JSDOM } = require('jsdom');
      const dom = new JSDOM(htmlContent);
      const document = dom.window.document;

      // Simulate the extraction logic
      const extractedData = this.extractPlannerData(document);
      this.results[viewType] = extractedData;

      console.log('✅ Plan Information:');
      console.log(`   Plan Name: ${extractedData.planData.planName || 'Not found'}`);
      console.log(`   Current View: ${extractedData.planData.currentView || 'Not found'}`);
      console.log(`   Buckets Found: ${extractedData.planData.buckets?.length || 0}`);

      console.log('\n📋 Task Information:');
      console.log(`   Total Tasks: ${extractedData.taskData.length}`);
      console.log(`   Tasks with Names: ${extractedData.taskData.filter(t => t.name).length}`);
      console.log(`   Tasks with Assignments: ${extractedData.taskData.filter(t => t.assignedTo).length}`);
      console.log(`   Tasks with Progress: ${extractedData.taskData.filter(t => t.progress !== undefined).length}`);

      if (extractedData.taskData.length > 0) {
        console.log('\n📝 Sample Tasks:');
        extractedData.taskData.slice(0, 3).forEach((task, index) => {
          console.log(`   ${index + 1}. ${task.name || 'Unnamed'}`);
          console.log(`      Assigned: ${task.assignedTo || 'Unassigned'}`);
          console.log(`      Progress: ${task.progress || 0}%`);
          console.log(`      Priority: ${task.priority || 'Not set'}`);
        });
      }

      console.log('\n');

    } catch (error) {
      console.error(`❌ Error testing ${filename}:`, error.message);
    }
  }

  extractPlannerData(document) {
    const planData = {};
    const taskData = [];

    // Extract plan info
    this.extractPlanInfo(document, planData);

    // Extract task info
    this.extractTaskInfo(document, taskData);

    // Extract bucket info
    this.extractBucketInfo(document, planData);

    return { planData, taskData };
  }

  extractPlanInfo(document, planData) {
    // Extract plan name from breadcrumb
    const planNameElement = document.querySelector('.ms-Breadcrumb-itemLink');
    if (planNameElement) {
      planData.planName = planNameElement.textContent.trim();
    }

    // Extract project title from page
    const projectTitleElements = document.querySelectorAll('[aria-label*="Project"]');
    projectTitleElements.forEach(el => {
      if (el.textContent && el.textContent.trim()) {
        planData.projectTitle = el.textContent.trim();
      }
    });

    // Extract current view
    const viewElements = document.querySelectorAll('[aria-label*="View"]');
    viewElements.forEach(el => {
      if (el.classList.contains('selected') || el.getAttribute('aria-selected') === 'true') {
        planData.currentView = el.textContent.trim();
      }
    });

    // Look for any element that might contain the plan name
    if (!planData.planName) {
      const titleElements = document.querySelectorAll('h1, h2, title, [class*="title"], [class*="plan"]');
      titleElements.forEach(el => {
        const text = el.textContent.trim();
        if (text && text.includes('Plan') && text.length < 100) {
          planData.planName = text;
        }
      });
    }
  }

  extractTaskInfo(document, taskData) {
    this.extractGridViewTasks(document, taskData);
    this.extractBoardViewTasks(document, taskData);
    this.extractTaskDetails(document, taskData);
  }

  extractGridViewTasks(document, taskData) {
    const taskRows = document.querySelectorAll('[role="row"]');

    taskRows.forEach((row, index) => {
      if (index === 0) return; // Skip header row

      const task = {};
      const cells = row.querySelectorAll('[role="gridcell"]');

      cells.forEach(cell => {
        const ariaLabel = cell.getAttribute('aria-label');
        if (!ariaLabel) return;

        // Extract task name
        if (ariaLabel.includes('Task Name')) {
          const taskNameElement = cell.querySelector('span, div, button');
          if (taskNameElement) {
            task.name = taskNameElement.textContent.trim();
          }
        }

        // Extract assigned user
        if (ariaLabel.includes('Assigned to')) {
          task.assignedTo = this.extractAssignedUser(cell);
        }

        // Extract progress percentage
        if (ariaLabel.includes('% complete')) {
          const percentMatch = ariaLabel.match(/(\d+)%/);
          if (percentMatch) {
            task.progress = parseInt(percentMatch[1]);
          }
        }

        // Extract priority
        if (ariaLabel.includes('Priority')) {
          const priorityMatch = ariaLabel.match(/Priority\s+(\w+)/);
          if (priorityMatch) {
            task.priority = priorityMatch[1];
          }
        }

        // Extract bucket
        if (ariaLabel.includes('Bucket')) {
          const bucketMatch = ariaLabel.match(/Bucket\s+([^.]+)/);
          if (bucketMatch) {
            task.bucket = bucketMatch[1].trim();
          }
        }

        // Extract completion status
        if (ariaLabel.includes('Mark as completed')) {
          task.completed = ariaLabel.includes('checked');
        }
      });

      if (task.name) {
        task.id = this.generateTaskId(task.name);
        task.source = 'grid';
        taskData.push(task);
      }
    });
  }

  extractBoardViewTasks(document, taskData) {
    const taskCards = document.querySelectorAll('[data-testid="task-card"], .task-card, [class*="task"], [class*="card"]');

    taskCards.forEach(card => {
      const task = {};

      const nameElement = card.querySelector('h3, h4, .task-title, [class*="title"]');
      if (nameElement) {
        task.name = nameElement.textContent.trim();
      }

      const avatars = card.querySelectorAll('.avatar, [class*="avatar"], [class*="assigned"]');
      if (avatars.length > 0) {
        task.assignedTo = Array.from(avatars).map(avatar =>
          avatar.getAttribute('title') || avatar.getAttribute('alt') || 'Unknown'
        );
      }

      if (task.name) {
        task.id = this.generateTaskId(task.name);
        task.source = 'board';
        taskData.push(task);
      }
    });
  }

  extractTaskDetails(document, taskData) {
    // Look for task names in various possible locations
    const taskSelectors = [
      '[aria-label*="Task Name"]',
      'button[aria-label*="Task"]',
      '[class*="task"] button',
      '[role="gridcell"] button'
    ];

    taskSelectors.forEach(selector => {
      const elements = document.querySelectorAll(selector);
      elements.forEach(el => {
        const ariaLabel = el.getAttribute('aria-label');
        if (ariaLabel) {
          // Try to extract task name from aria-label
          const taskNameMatch = ariaLabel.match(/Task Name\s+([^.]+)/);
          if (taskNameMatch) {
            const task = {
              name: taskNameMatch[1].trim(),
              id: this.generateTaskId(taskNameMatch[1].trim()),
              source: 'aria-label'
            };

            // Check for additional info in the aria-label
            if (ariaLabel.includes('completed')) {
              task.completed = true;
            }

            taskData.push(task);
          }
        }
      });
    });
  }

  extractBucketInfo(document, planData) {
    const buckets = [];
    const bucketHeaders = document.querySelectorAll('[class*="bucket"], [class*="column"] h2, [class*="column"] h3');

    bucketHeaders.forEach(header => {
      const bucketName = header.textContent.trim();
      if (bucketName && !buckets.includes(bucketName)) {
        buckets.push(bucketName);
      }
    });

    // Also look for bucket mentions in aria-labels
    const bucketElements = document.querySelectorAll('[aria-label*="Bucket"]');
    bucketElements.forEach(el => {
      const ariaLabel = el.getAttribute('aria-label');
      const bucketMatch = ariaLabel.match(/Bucket\s+([^.]+)/);
      if (bucketMatch) {
        const bucketName = bucketMatch[1].trim();
        if (!buckets.includes(bucketName)) {
          buckets.push(bucketName);
        }
      }
    });

    planData.buckets = buckets;
  }

  extractAssignedUser(element) {
    const userElement = element.querySelector('[title], [alt], span, div');
    if (userElement) {
      return userElement.getAttribute('title') ||
             userElement.getAttribute('alt') ||
             userElement.textContent.trim();
    }
    return null;
  }

  generateTaskId(taskName) {
    return taskName.toLowerCase().replace(/[^a-z0-9]/g, '-').substring(0, 50);
  }

  generateReport() {
    console.log('\n📊 EXTRACTION REPORT');
    console.log('='.repeat(50));

    const entries = Object.entries(this.results);

    console.log('\n🔍 Pattern Analysis:');
    entries.forEach(([viewType, data]) => {
      console.log(`${viewType} - Tasks Found: ${data.taskData?.length || 0}`);
    });

    const primary = entries[0]?.[1] || {};

    console.log('\n📈 Success Metrics:');
    console.log(`Plan Name Detection: ${primary.planData?.planName ? '✅' : '❌'}`);
    console.log(`Task Extraction: ${(primary.taskData?.length || 0) > 0 ? '✅' : '❌'}`);
    console.log(`Bucket Detection: ${(primary.planData?.buckets?.length || 0) > 0 ? '✅' : '❌'}`);

    console.log('\n🔧 Recommendations:');
    if ((general.taskData?.length || 0) === 0) {
      console.log('- Review task extraction selectors');
      console.log('- Check for different DOM structure in this Planner version');
    }

    if (!general.planData?.planName) {
      console.log('- Review plan name extraction logic');
      console.log('- Check breadcrumb structure');
    }

    console.log('\n✨ Extraction Complete!\n');
  }
}

// Run the test if this file is executed directly
if (require.main === module) {
  // Check if JSDOM is available
  try {
    require('jsdom');
    const tester = new PlannerTestExtractor();
    tester.testBothFiles().catch(console.error);
  } catch (error) {
    console.log('⚠️  JSDOM not available. Install with: npm install jsdom');
    console.log('Running basic HTML analysis instead...\n');

    // Fallback to basic analysis
    const fs = require('fs');

    const fallbackTargets = process.argv.slice(2);
    const targets = fallbackTargets.length > 0
      ? fallbackTargets
      : [path.join('Tests', 'multi task plan.html')];

    targets.forEach(filename => {
      console.log(`📄 Analyzing ${filename}:`);
      try {
        const content = fs.readFileSync(filename, 'utf8');
        console.log(`   File size: ${(content.length / 1024 / 1024).toFixed(2)} MB`);
        console.log(`   Contains "Task Name": ${content.includes('Task Name') ? '✅' : '❌'}`);
        console.log(`   Contains "Assigned to": ${content.includes('Assigned to') ? '✅' : '❌'}`);
        console.log(`   Contains "% complete": ${content.includes('% complete') ? '✅' : '❌'}`);
        console.log(`   Contains role="row": ${content.includes('role="row"') ? '✅' : '❌'}`);
        console.log(`   Contains role="gridcell": ${content.includes('role="gridcell"') ? '✅' : '❌'}`);
      } catch (error) {
        console.log(`   ❌ Error reading file: ${error.message}`);
      }
      console.log('');
    });
  }
}

module.exports = PlannerTestExtractor;
