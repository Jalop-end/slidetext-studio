const { execSync } = require('child_process');

async function main() {
  const credOut = execSync('git credential fill', { input: 'protocol=https\nhost=github.com\n' }).toString();
  const token = credOut.match(/password=(.+)/)[1].trim();
  const username = credOut.match(/username=(.+)/)[1].trim();

  const repoName = 'slidetext-studio';
  console.log(`Checking / creating repository ${username}/${repoName} on GitHub...`);

  // Check if repo already exists
  const checkRes = await fetch(`https://api.github.com/repos/${username}/${repoName}`, {
    headers: {
      Authorization: `Bearer ${token}`,
      'User-Agent': 'Node-App',
      Accept: 'application/vnd.github+json',
    },
  });

  if (checkRes.status === 404) {
    console.log(`Repository does not exist yet. Creating repository ${repoName}...`);
    const createRes = await fetch('https://api.github.com/user/repos', {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'User-Agent': 'Node-App',
        Accept: 'application/vnd.github+json',
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        name: repoName,
        description: 'SlideText Studio - PowerPoint Editor & Fast Free AI Translator (100% layout & visual preservation)',
        private: false,
        auto_init: false,
      }),
    });

    if (!createRes.ok) {
      const err = await createRes.text();
      throw new Error(`Failed to create repository on GitHub: ${err}`);
    }
    const createdData = await createRes.json();
    console.log(`✅ GitHub repository created: ${createdData.html_url}`);
  } else if (checkRes.ok) {
    const existingData = await checkRes.json();
    console.log(`ℹ️ GitHub repository already exists: ${existingData.html_url}`);
  } else {
    const err = await checkRes.text();
    throw new Error(`GitHub API check error: ${err}`);
  }

  // Initialize git repo locally
  console.log('📦 Initializing local git repository...');
  try {
    execSync('git init', { stdio: 'inherit' });
    execSync('git config user.name "Jalop-end"', { stdio: 'inherit' });
    execSync('git config user.email "181258586+Jalop-end@users.noreply.github.com"', { stdio: 'inherit' });
    execSync('git branch -M main', { stdio: 'inherit' });

    console.log('📝 Staging files...');
    execSync('git add .', { stdio: 'inherit' });

    console.log('💾 Creating commit...');
    execSync('git commit -m "Initial commit: SlideText Studio - PowerPoint Editor & AI Translator"', { stdio: 'inherit' });

    const remoteUrl = `https://github.com/${username}/${repoName}.git`;
    console.log(`🔗 Configuring remote origin: ${remoteUrl}...`);
    try {
      execSync(`git remote add origin ${remoteUrl}`, { stdio: 'inherit' });
    } catch (_) {
      execSync(`git remote set-url origin ${remoteUrl}`, { stdio: 'inherit' });
    }

    console.log('🚀 Pushing to GitHub (main)...');
    execSync('git push -u origin main --force', { stdio: 'inherit' });

    console.log(`\n🎉 HOTOVO! Váš projekt byl úspěšně nahrán na GitHub: https://github.com/${username}/${repoName}`);
  } catch (err) {
    console.error('Git error:', err);
    process.exit(1);
  }
}

main().catch(err => {
  console.error('Push failed:', err);
  process.exit(1);
});
