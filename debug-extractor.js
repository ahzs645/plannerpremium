/**
 * Debug script to inspect saved Planner HTML pages without relying on sample data
 */

const fs = require('fs');
const path = require('path');

const defaultPath = path.join(__dirname, 'Tests', 'multi task plan.html');
const htmlPath = process.argv[2] || defaultPath;

try {
  const htmlContent = fs.readFileSync(htmlPath, 'utf8');
  console.log(`📄 Loaded: ${htmlPath}`);
  console.log(`Size: ${(htmlContent.length / 1024 / 1024).toFixed(2)} MB`);
  console.log('-'.repeat(60));

  // Plan title patterns
  const titleMatch = htmlContent.match(/<title>([^<]*)<\/title>/i);
  if (titleMatch) {
    console.log(`Page <title>: ${titleMatch[1].trim()}`);
  }

  const headingRegex = /role="heading"[^>]*>([^<]+)<\/[^>]+>/gi;
  const headings = new Set();
  let headingMatch;
  while ((headingMatch = headingRegex.exec(htmlContent)) !== null) {
    const text = headingMatch[1].trim();
    if (text) headings.add(text);
  }
  console.log(`Headings detected: ${headings.size}`);

  // Task extraction preview via aria-labels
  const taskRegex = /aria-label="([^"]*Task Name[^"]*)"/g;
  const taskNames = new Map();
  let taskMatch;
  while ((taskMatch = taskRegex.exec(htmlContent)) !== null) {
    const label = taskMatch[1];
    const nameMatch = label.match(/Task Name\s+([^\.]+)/);
    if (!nameMatch) continue;
    const name = nameMatch[1].trim();
    const count = taskNames.get(name) || 0;
    taskNames.set(name, count + 1);
  }

  console.log(`Tasks detected from aria-labels: ${taskNames.size}`);
  for (const [name, count] of taskNames) {
    console.log(`  - ${name} (${count})`);
  }

  // Assignment preview
  const assignmentRegex = /aria-label="([^"]*Assigned to[^"]*)"/gi;
  const assignments = new Map();
  let assignmentMatch;
  while ((assignmentMatch = assignmentRegex.exec(htmlContent)) !== null) {
    const label = assignmentMatch[1];
    const assigneeMatch = label.match(/Assigned to\s+([^\.]+)/i);
    if (!assigneeMatch) continue;
    const assignee = assigneeMatch[1].trim();
    const count = assignments.get(assignee) || 0;
    assignments.set(assignee, count + 1);
  }

  if (assignments.size > 0) {
    console.log('Assignments detected:');
    for (const [assignee, count] of assignments) {
      console.log(`  - ${assignee} (${count})`);
    }
  }

  // Structural cues
  const roleRowCount = (htmlContent.match(/role="row"/g) || []).length;
  const roleGridcellCount = (htmlContent.match(/role="gridcell"/g) || []).length;
  const checkboxCount = (htmlContent.match(/type="checkbox"/g) || []).length;
  console.log(`role="row": ${roleRowCount}`);
  console.log(`role="gridcell": ${roleGridcellCount}`);
  console.log(`checkbox inputs: ${checkboxCount}`);

  console.log('-'.repeat(60));
  console.log('✅ Analysis complete');

} catch (error) {
  console.error(`❌ Unable to read ${htmlPath}:`, error.message);
  const testsDir = path.join(__dirname, 'Tests');
  if (fs.existsSync(testsDir)) {
    console.log('\nAvailable HTML snapshots:');
    fs.readdirSync(testsDir)
      .filter(file => file.toLowerCase().endsWith('.html'))
      .forEach(file => console.log(`  - ${path.join('Tests', file)}`));
  }
}
