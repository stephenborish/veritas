const express = require('express');
const app = express();
const path = require('path');
const fs = require('fs');

// We need to serve the HTML file but also mock out `google.script.run` inside it.
app.get('/', (req, res) => {
    let html = fs.readFileSync(path.join(__dirname, 'TeacherApp.html'), 'utf8');

    // Inject google.script.run mock
    const mock = `
    <script>
      window.google = {
        script: {
          run: {
            withSuccessHandler: function(cb) {
              const runObj = {
                withFailureHandler: function() {
                  return runObj;
                },
                getActiveSession: () => setTimeout(() => cb(null), 10),
                getQuestionSets: () => setTimeout(() => cb([{id: 'qs1', name: 'Biology Test', questionCount: 10}]), 10),
                getArchivedSessions: () => setTimeout(() => cb([]), 10),
                getCourses: () => setTimeout(() => cb([{id: 'c1', name: 'AP Biology'}, {id: 'c2', name: 'Integrated Science 1'}]), 10),
                getRosters: () => setTimeout(() => cb({
                    'c1_1': {courseId: 'c1', block: '1', students: []},
                    'c2_1': {courseId: 'c2', block: '1', students: []}
                }), 10),
                activateSession: (cfg) => {
                    console.log("ACTIVATE CALLED", cfg);
                    window.lastCfg = cfg;
                    setTimeout(() => cb({sessionId: 'sess1', code: 'ABC12'}), 10);
                }
              };
              return runObj;
            }
          }
        }
      };
    </script>
    `;

    html = html.replace('<head>', '<head>' + mock);
    res.send(html);
});

app.listen(8001, () => {
    console.log('Server running on port 8001');
});