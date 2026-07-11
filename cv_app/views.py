import re
from copy import deepcopy
from io import BytesIO

from django.contrib import messages
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required
from django.forms import modelformset_factory
from django.http import HttpResponse
from django.shortcuts import redirect, render
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import ListFlowable, ListItem, Paragraph, SimpleDocTemplate, Spacer

from .forms import (
    BenefitItemForm,
    CertificationForm,
    EducationForm,
    ExperienceForm,
    FAQItemForm,
    FeatureItemForm,
    HonorAwardForm,
    LandingPageContentForm,
    ProcessStepForm,
    ProjectForm,
    SkillCategoryForm,
    TestimonialForm,
)
from .models import (
    BenefitItem,
    FAQItem,
    FeatureItem,
    LandingPageContent,
    ProcessStep,
    Testimonial,
)

EMPTY_BUILDER_DATA = {
    'personal': {
        'full_name': '',
        'email': '',
        'phone': '',
        'linkedin': '',
        'github': '',
        'portfolio': '',
        'location': '',
    },
    'job_requirements': {
        'job_title': '',
        'role': '',
        'job_description': '',
        'keywords': '',
        'required_skills': '',
    },
    'skill_categories': [],
    'projects': [],
    'education': [],
    'experience': [],
    'certifications': [],
    'honors_awards': [],
}

DEFAULT_SKILL_CATEGORIES = [
    'Programming Languages',
    'Machine Learning & AI',
    'Data Analysis & Visualization',
    'Database Management',
    'Backend Development',
    'Frontend Development',
    'Cloud & Deployment',
    'DevOps & Infrastructure',
    'Web Scraping & Automation',
    'Dev Tools & IDEs',
    'Software Testing & QA',
    'Project Management & Soft Skills',
    'Mobile Development',
    'Cybersecurity',
    'Data Engineering',
    'API Development & Integration',
    'UI/UX Design',
    'Game Development',
    'Blockchain & Web3',
    'Embedded Systems & IoT',
]

SKILL_KEYWORD_MAP = {
    'Programming Languages': [
        ('python', 'Python'),
        ('c++', 'C++'),
        ('c', 'C'),
        ('java', 'Java'),
        ('javascript', 'JavaScript'),
        (' js', 'JavaScript'),
        ('typescript', 'TypeScript'),
        ('php', 'PHP'),
        ('ruby', 'Ruby'),
        ('go', 'Go'),
        ('golang', 'Go'),
        ('kotlin', 'Kotlin'),
        ('swift', 'Swift'),
        ('sql', 'SQL'),
        ('bash', 'Bash'),
        ('shell', 'Shell'),
        ('r', 'R'),
        ('scala', 'Scala'),
        ('rust', 'Rust'),
        ('matlab', 'MATLAB'),
        ('dart', 'Dart'),
        ('lua', 'Lua'),
        ('perl', 'Perl'),
        ('objective-c', 'Objective-C'),
        ('c#', 'C#'),
        ('dotnet', '.NET'),
        ('vb.net', 'VB.NET'),
        ('f#', 'F#'),
        ('haskell', 'Haskell'),
        ('clojure', 'Clojure'),
        ('elixir', 'Elixir'),
        ('erlang', 'Erlang'),
    ],
    'Machine Learning & AI': [
        ('machine learning', 'Machine Learning'),
        ('ml', 'Machine Learning'),
        ('deep learning', 'Deep Learning'),
        ('artificial intelligence', 'Artificial Intelligence'),
        (' ai ', 'Artificial Intelligence'),
        ('scikit-learn', 'Scikit-learn'),
        ('sklearn', 'Scikit-learn'),
        ('tensorflow', 'TensorFlow'),
        ('pytorch', 'PyTorch'),
        ('keras', 'Keras'),
        ('feature engineering', 'Feature Engineering'),
        ('data preprocessing', 'Data Preprocessing'),
        ('feature scaling', 'Feature Scaling and Normalization'),
        ('normalization', 'Feature Scaling and Normalization'),
        ('hyperparameter tuning', 'Hyperparameter Tuning'),
        ('model validation', 'Model Validation'),
        ('mlops', 'MLOps'),
        ('ci/cd', 'CI / CD'),
        ('computer vision', 'Computer Vision'),
        ('cv', 'Computer Vision'),
        ('nlp', 'NLP'),
        ('natural language processing', 'NLP'),
        ('reinforcement learning', 'Reinforcement Learning'),
        ('xgboost', 'XGBoost'),
        ('lightgbm', 'LightGBM'),
        ('optuna', 'Optuna'),
        ('docker', 'Docker'),
        ('hugging face', 'Hugging Face'),
        ('transformers', 'Transformers'),
        ('bert', 'BERT'),
        ('gpt', 'GPT'),
        ('llm', 'Large Language Models'),
        ('large language model', 'Large Language Models'),
        ('generative ai', 'Generative AI'),
        ('rag', 'RAG'),
        ('retrieval augmented generation', 'RAG'),
        ('langchain', 'LangChain'),
        ('llama', 'LLaMA'),
        ('stable diffusion', 'Stable Diffusion'),
        ('gan', 'GAN'),
        ('generative adversarial network', 'GAN'),
        ('yolo', 'YOLO'),
        ('opencv', 'OpenCV'),
        ('mediapipe', 'MediaPipe'),
        ('onnx', 'ONNX'),
        ('model deployment', 'Model Deployment'),
        ('model serving', 'Model Serving'),
        ('tensorrt', 'TensorRT'),
        ('triton', 'Triton Inference Server'),
    ],
    'Data Analysis & Visualization': [
        ('pandas', 'Pandas'),
        ('numpy', 'NumPy'),
        ('matplotlib', 'Matplotlib'),
        ('seaborn', 'Seaborn'),
        ('tableau', 'Tableau'),
        ('power bi', 'Power BI'),
        ('powerbi', 'Power BI'),
        ('metabase', 'Metabase'),
        ('eda', 'EDA'),
        ('exploratory data analysis', 'EDA'),
        ('data analysis', 'Data Analysis'),
        ('analytics', 'Analytics'),
        ('visualization', 'Data Visualization'),
        ('plotly', 'Plotly'),
        ('ggplot', 'ggplot'),
        ('superset', 'Apache Superset'),
        ('looker', 'Looker'),
        ('redash', 'Redash'),
        ('grafana', 'Grafana'),
        ('kibana', 'Kibana'),
        ('excel', 'Excel'),
        ('spreadsheet', 'Spreadsheets'),
        ('pivot table', 'Pivot Tables'),
        ('vlookup', 'VLOOKUP'),
        ('dax', 'DAX'),
        ('power query', 'Power Query'),
        ('statistics', 'Statistics'),
        ('statistical analysis', 'Statistical Analysis'),
        ('hypothesis testing', 'Hypothesis Testing'),
        ('regression', 'Regression Analysis'),
        ('time series', 'Time Series Analysis'),
        ('forecasting', 'Forecasting'),
        ('cohort analysis', 'Cohort Analysis'),
        ('funnel analysis', 'Funnel Analysis'),
    ],
    'Database Management': [
        ('postgresql', 'PostgreSQL'),
        ('postgres', 'PostgreSQL'),
        ('mysql', 'MySQL'),
        ('mongodb', 'MongoDB'),
        ('mongo', 'MongoDB'),
        ('sqlite', 'SQLite'),
        ('nosql', 'NoSQL'),
        ('database', 'Databases'),
        ('sql', 'SQL'),
        ('redis', 'Redis'),
        ('cassandra', 'Cassandra'),
        ('dynamodb', 'DynamoDB'),
        ('oracle', 'Oracle'),
        ('snowflake', 'Snowflake'),
        ('bigquery', 'BigQuery'),
        ('redshift', 'Redshift'),
        ('clickhouse', 'ClickHouse'),
        ('elasticsearch', 'Elasticsearch'),
        ('solr', 'Solr'),
        ('neo4j', 'Neo4j'),
        ('graph database', 'Graph Database'),
        ('timeseries', 'Time Series Database'),
        ('influxdb', 'InfluxDB'),
        ('couchdb', 'CouchDB'),
        ('couchbase', 'Couchbase'),
        ('mariadb', 'MariaDB'),
        ('sql server', 'SQL Server'),
        ('tsql', 'T-SQL'),
        ('pl/sql', 'PL/SQL'),
        ('stored procedure', 'Stored Procedures'),
        ('database design', 'Database Design'),
        ('data modeling', 'Data Modeling'),
        ('er diagram', 'ER Diagram'),
        ('normalization', 'Database Normalization'),
        ('indexing', 'Database Indexing'),
        ('query optimization', 'Query Optimization'),
        ('acid', 'ACID Transactions'),
        ('replication', 'Database Replication'),
        ('sharding', 'Database Sharding'),
    ],
    'Backend Development': [
        ('django', 'Django'),
        ('drf', 'Django REST Framework'),
        ('django rest', 'Django REST Framework'),
        ('flask', 'Flask'),
        ('fastapi', 'FastAPI'),
        ('rest api', 'REST APIs'),
        ('restful', 'REST APIs'),
        ('api', 'API Development'),
        ('backend', 'Backend Development'),
        ('authentication', 'Authentication'),
        ('oauth', 'OAuth'),
        ('oauth2', 'OAuth2'),
        ('jwt', 'JWT'),
        ('json web token', 'JWT'),
        ('session', 'Session Management'),
        ('cookie', 'Cookie Management'),
        ('node.js', 'Node.js'),
        ('nodejs', 'Node.js'),
        ('express', 'Express.js'),
        ('expressjs', 'Express.js'),
        ('nestjs', 'NestJS'),
        ('koa', 'Koa'),
        ('spring boot', 'Spring Boot'),
        ('spring', 'Spring Framework'),
        ('graphql', 'GraphQL'),
        ('grpc', 'gRPC'),
        ('protocol buffers', 'Protocol Buffers'),
        ('protobuf', 'Protocol Buffers'),
        ('microservices', 'Microservices'),
        ('service mesh', 'Service Mesh'),
        ('istio', 'Istio'),
        ('linkerd', 'Linkerd'),
        ('message queue', 'Message Queues'),
        ('rabbitmq', 'RabbitMQ'),
        ('kafka', 'Apache Kafka'),
        ('pub/sub', 'Pub/Sub'),
        ('event driven', 'Event-Driven Architecture'),
        ('cqrs', 'CQRS'),
        ('event sourcing', 'Event Sourcing'),
        ('golang', 'Go'),
        ('gin', 'Gin'),
        ('echo', 'Echo'),
        ('fiber', 'Fiber'),
        ('laravel', 'Laravel'),
        ('symfony', 'Symfony'),
        ('codeigniter', 'CodeIgniter'),
        ('asp.net', 'ASP.NET Core'),
        ('dotnet core', 'ASP.NET Core'),
        ('entity framework', 'Entity Framework'),
        ('ruby on rails', 'Ruby on Rails'),
        ('rails', 'Ruby on Rails'),
    ],
    'Frontend Development': [
        ('react', 'React'),
        ('reactjs', 'React'),
        ('react.js', 'React'),
        ('next.js', 'Next.js'),
        ('nextjs', 'Next.js'),
        ('vue', 'Vue.js'),
        ('vuejs', 'Vue.js'),
        ('vue.js', 'Vue.js'),
        ('nuxt', 'Nuxt.js'),
        ('nuxtjs', 'Nuxt.js'),
        ('angular', 'Angular'),
        ('angularjs', 'AngularJS'),
        ('svelte', 'Svelte'),
        ('sveltekit', 'SvelteKit'),
        ('html', 'HTML'),
        ('css', 'CSS'),
        ('scss', 'SCSS'),
        ('sass', 'Sass'),
        ('less', 'Less'),
        ('tailwind', 'Tailwind CSS'),
        ('tailwindcss', 'Tailwind CSS'),
        ('bootstrap', 'Bootstrap'),
        ('material ui', 'Material UI'),
        ('mui', 'MUI'),
        ('chakra ui', 'Chakra UI'),
        ('ant design', 'Ant Design'),
        ('jest', 'Jest'),
        ('vitest', 'Vitest'),
        ('cypress', 'Cypress'),
        ('playwright', 'Playwright'),
        ('storybook', 'Storybook'),
        ('webpack', 'Webpack'),
        ('vite', 'Vite'),
        ('rollup', 'Rollup'),
        ('esbuild', 'esbuild'),
        ('parcel', 'Parcel'),
        ('babel', 'Babel'),
        ('typescript', 'TypeScript'),
        ('ts', 'TypeScript'),
        ('eslint', 'ESLint'),
        ('prettier', 'Prettier'),
        ('responsive design', 'Responsive Design'),
        ('web accessibility', 'Web Accessibility'),
        ('a11y', 'Accessibility'),
        ('wcag', 'WCAG'),
        ('pwa', 'Progressive Web Apps'),
        ('service worker', 'Service Workers'),
        ('web components', 'Web Components'),
        ('shadow dom', 'Shadow DOM'),
        ('canvas', 'Canvas API'),
        ('webgl', 'WebGL'),
        ('three.js', 'Three.js'),
        ('gsap', 'GSAP'),
        ('framer motion', 'Framer Motion'),
    ],
    'Cloud & Deployment': [
        ('aws', 'AWS'),
        ('amazon web services', 'AWS'),
        ('ec2', 'EC2'),
        ('s3', 'S3'),
        ('lambda', 'Lambda'),
        ('rds', 'RDS'),
        ('dynamodb', 'DynamoDB'),
        ('cloudfront', 'CloudFront'),
        ('cloudformation', 'CloudFormation'),
        ('cdk', 'AWS CDK'),
        ('gcp', 'Google Cloud Platform'),
        ('google cloud', 'Google Cloud Platform'),
        ('compute engine', 'Compute Engine'),
        ('cloud run', 'Cloud Run'),
        ('cloud functions', 'Cloud Functions'),
        ('firestore', 'Firestore'),
        ('pubsub', 'Pub/Sub'),
        ('bigquery', 'BigQuery'),
        ('azure', 'Azure'),
        ('microsoft azure', 'Azure'),
        ('azure functions', 'Azure Functions'),
        ('aks', 'AKS'),
        ('docker', 'Docker'),
        ('container', 'Containerization'),
        ('kubernetes', 'Kubernetes'),
        ('k8s', 'Kubernetes'),
        ('helm', 'Helm'),
        ('ci/cd', 'CI / CD'),
        ('continuous integration', 'Continuous Integration'),
        ('continuous deployment', 'Continuous Deployment'),
        ('deployment', 'Deployment'),
        ('render', 'Render'),
        ('vercel', 'Vercel'),
        ('netlify', 'Netlify'),
        ('supabase', 'Supabase'),
        ('mongodb atlas', 'MongoDB Atlas'),
        ('planetscale', 'PlanetScale'),
        ('neon', 'Neon'),
        ('railway', 'Railway'),
        ('fly.io', 'Fly.io'),
        ('heroku', 'Heroku'),
        ('digitalocean', 'DigitalOcean'),
        ('linode', 'Linode'),
        ('vultr', 'Vultr'),
        ('cloud', 'Cloud Computing'),
        ('terraform', 'Terraform'),
        ('infrastructure as code', 'Infrastructure as Code'),
        ('iac', 'Infrastructure as Code'),
        ('ansible', 'Ansible'),
        ('jenkins', 'Jenkins'),
        ('github actions', 'GitHub Actions'),
        ('gitlab ci', 'GitLab CI'),
        ('circleci', 'CircleCI'),
        ('travis ci', 'Travis CI'),
        ('argo cd', 'Argo CD'),
        ('flux', 'Flux'),
        ('prometheus', 'Prometheus'),
        ('grafana', 'Grafana'),
        ('datadog', 'Datadog'),
        ('new relic', 'New Relic'),
        ('sentry', 'Sentry'),
        ('elk stack', 'ELK Stack'),
        ('elastic stack', 'Elastic Stack'),
        ('loki', 'Loki'),
        ('tempo', 'Tempo'),
        ('jaeger', 'Jaeger'),
        ('zipkin', 'Zipkin'),
    ],
    'DevOps & Infrastructure': [
        ('devops', 'DevOps'),
        ('ci/cd', 'CI/CD'),
        ('pipeline', 'CI/CD Pipelines'),
        ('jenkins', 'Jenkins'),
        ('github actions', 'GitHub Actions'),
        ('gitlab ci', 'GitLab CI'),
        ('circleci', 'CircleCI'),
        ('travis', 'Travis CI'),
        ('argo', 'Argo'),
        ('tekton', 'Tekton'),
        ('spinnaker', 'Spinnaker'),
        ('terraform', 'Terraform'),
        ('pulumi', 'Pulumi'),
        ('ansible', 'Ansible'),
        ('chef', 'Chef'),
        ('puppet', 'Puppet'),
        ('saltstack', 'SaltStack'),
        ('kubernetes', 'Kubernetes'),
        ('k8s', 'Kubernetes'),
        ('helm', 'Helm'),
        ('kustomize', 'Kustomize'),
        ('operator', 'Kubernetes Operators'),
        ('istio', 'Istio'),
        ('linkerd', 'Linkerd'),
        ('consul', 'Consul'),
        ('vault', 'HashiCorp Vault'),
        ('nomad', 'Nomad'),
        ('docker', 'Docker'),
        ('containerd', 'containerd'),
        ('podman', 'Podman'),
        ('buildah', 'Buildah'),
        ('kaniko', 'Kaniko'),
        ('skaffold', 'Skaffold'),
        ('tilt', 'Tilt'),
        ('prometheus', 'Prometheus'),
        ('grafana', 'Grafana'),
        ('alertmanager', 'Alertmanager'),
        ('loki', 'Loki'),
        ('tempo', 'Tempo'),
        ('jaeger', 'Jaeger'),
        ('opentelemetry', 'OpenTelemetry'),
        ('elk', 'ELK Stack'),
        ('efk', 'EFK Stack'),
        ('datadog', 'Datadog'),
        ('new relic', 'New Relic'),
        ('sentry', 'Sentry'),
        ('pagerduty', 'PagerDuty'),
        ('opsgenie', 'Opsgenie'),
        ('linux', 'Linux'),
        ('bash', 'Bash Scripting'),
        ('shell scripting', 'Shell Scripting'),
        ('systemd', 'systemd'),
        ('nginx', 'NGINX'),
        ('apache', 'Apache HTTP Server'),
        ('haproxy', 'HAProxy'),
        ('traefik', 'Traefik'),
        ('envoy', 'Envoy Proxy'),
        ('dns', 'DNS'),
        ('ssl', 'SSL/TLS'),
        ('tls', 'SSL/TLS'),
        ('letsencrypt', 'Let\'s Encrypt'),
        ('cert-manager', 'cert-manager'),
    ],
    'Web Scraping & Automation': [
        ('web scraping', 'Web Scraping'),
        ('scraping', 'Web Scraping'),
        ('requests', 'Requests'),
        ('httpx', 'HTTPX'),
        ('aiohttp', 'aiohttp'),
        ('selenium', 'Selenium'),
        ('playwright', 'Playwright'),
        ('puppeteer', 'Puppeteer'),
        ('beautifulsoup', 'BeautifulSoup'),
        ('bs4', 'BeautifulSoup'),
        ('lxml', 'lxml'),
        ('scrapy', 'Scrapy'),
        ('automation', 'Automation'),
        ('bots', 'Automation Bots'),
        ('rpa', 'RPA'),
        ('robotic process automation', 'RPA'),
        ('zapier', 'Zapier'),
        ('n8n', 'n8n'),
        ('make', 'Make (Integromat)'),
        ('ifttt', 'IFTTT'),
        ('workflow automation', 'Workflow Automation'),
    ],
    'Dev Tools & IDEs': [
        ('git', 'Git'),
        ('github', 'GitHub'),
        ('gitlab', 'GitLab'),
        ('bitbucket', 'Bitbucket'),
        ('svn', 'SVN'),
        ('mercurial', 'Mercurial'),
        ('anaconda', 'Anaconda Navigator'),
        ('miniconda', 'Miniconda'),
        ('jupyter', 'Jupyter Notebook'),
        ('jupyterlab', 'JupyterLab'),
        ('colab', 'Google Colab'),
        ('kaggle', 'Kaggle'),
        ('visual studio', 'Visual Studio'),
        ('vscode', 'VS Code'),
        ('visual studio code', 'VS Code'),
        ('pycharm', 'PyCharm'),
        ('intellij', 'IntelliJ IDEA'),
        ('webstorm', 'WebStorm'),
        ('goland', 'GoLand'),
        ('rider', 'Rider'),
        ('clion', 'CLion'),
        ('phpstorm', 'PhpStorm'),
        ('datagrip', 'DataGrip'),
        ('vim', 'Vim'),
        ('neovim', 'Neovim'),
        ('emacs', 'Emacs'),
        ('sublime text', 'Sublime Text'),
        ('atom', 'Atom'),
        ('zed', 'Zed'),
        ('cursor', 'Cursor'),
        ('windsurf', 'Windsurf'),
        ('docker', 'Docker'),
        ('docker compose', 'Docker Compose'),
        ('podman', 'Podman'),
        ('make', 'Make'),
        ('cmake', 'CMake'),
        ('bazel', 'Bazel'),
        ('gradle', 'Gradle'),
        ('maven', 'Maven'),
        ('npm', 'npm'),
        ('yarn', 'Yarn'),
        ('pnpm', 'pnpm'),
        ('pip', 'pip'),
        ('poetry', 'Poetry'),
        ('pipenv', 'Pipenv'),
        ('conda', 'Conda'),
        ('uv', 'uv'),
        ('cargo', 'Cargo'),
        ('rustup', 'rustup'),
        ('nvm', 'nvm'),
        ('fnm', 'fnm'),
        ('asdf', 'asdf'),
        ('mise', 'mise'),
    ],
    'Software Testing & QA': [
        ('testing', 'Testing Fundamentals'),
        ('qa', 'Software Testing'),
        ('quality assurance', 'Quality Assurance'),
        ('pytest', 'Pytest'),
        ('unittest', 'unittest'),
        ('doctest', 'doctest'),
        ('nose', 'nose'),
        ('manual testing', 'Manual Testing'),
        ('automated testing', 'Automated Testing'),
        ('unit testing', 'Unit Testing'),
        ('integration testing', 'Integration Testing'),
        ('end to end testing', 'End-to-End Testing'),
        ('e2e testing', 'End-to-End Testing'),
        ('bug reporting', 'Bug Reporting'),
        ('test case', 'Test Case Design'),
        ('test plan', 'Test Planning'),
        ('selenium', 'Selenium Testing'),
        ('cypress', 'Cypress'),
        ('playwright', 'Playwright Testing'),
        ('puppeteer', 'Puppeteer Testing'),
        ('jest', 'Jest'),
        ('vitest', 'Vitest'),
        ('mocha', 'Mocha'),
        ('chai', 'Chai'),
        ('jasmine', 'Jasmine'),
        ('karma', 'Karma'),
        ('testcafe', 'TestCafe'),
        ('webdriverio', 'WebdriverIO'),
        ('nightwatch', 'Nightwatch'),
        ('codeceptjs', 'CodeceptJS'),
        ('robot framework', 'Robot Framework'),
        ('behave', 'Behave'),
        ('lettuce', 'Lettuce'),
        ('cucumber', 'Cucumber'),
        ('gherkin', 'Gherkin'),
        ('bdd', 'BDD'),
        ('tdd', 'TDD'),
        ('test driven development', 'TDD'),
        ('coverage', 'Code Coverage'),
        ('mutation testing', 'Mutation Testing'),
        ('property based testing', 'Property-Based Testing'),
        ('hypothesis', 'Hypothesis'),
        ('fuzzing', 'Fuzzing'),
        ('load testing', 'Load Testing'),
        ('performance testing', 'Performance Testing'),
        ('stress testing', 'Stress Testing'),
        ('jmeter', 'JMeter'),
        ('k6', 'k6'),
        ('locust', 'Locust'),
        ('gatling', 'Gatling'),
        ('vegeta', 'Vegeta'),
        ('contract testing', 'Contract Testing'),
        ('pact', 'Pact'),
        ('chaos engineering', 'Chaos Engineering'),
        ('litmus', 'Litmus'),
        ('gremlin', 'Gremlin'),
    ],
    'Project Management & Soft Skills': [
        ('agile', 'Agile'),
        ('scrum', 'Scrum'),
        ('kanban', 'Kanban'),
        ('jira', 'Jira'),
        ('confluence', 'Confluence'),
        ('trello', 'Trello'),
        ('notion', 'Notion'),
        ('asana', 'Asana'),
        ('monday.com', 'monday.com'),
        ('clickup', 'ClickUp'),
        ('linear', 'Linear'),
        ('github projects', 'GitHub Projects'),
        ('gitlab issues', 'GitLab Issues'),
        ('leadership', 'Leadership'),
        ('teamwork', 'Teamwork'),
        ('communication', 'Communication'),
        ('problem solving', 'Problem Solving'),
        ('documentation', 'Documentation'),
        ('mentoring', 'Mentoring'),
        ('coaching', 'Coaching'),
        ('code review', 'Code Review'),
        ('pair programming', 'Pair Programming'),
        ('technical writing', 'Technical Writing'),
        ('public speaking', 'Public Speaking'),
        ('presentation', 'Presentation Skills'),
        ('negotiation', 'Negotiation'),
        ('conflict resolution', 'Conflict Resolution'),
        ('time management', 'Time Management'),
        ('prioritization', 'Prioritization'),
        ('roadmap', 'Product Roadmap'),
        ('backlog', 'Backlog Management'),
        ('sprint planning', 'Sprint Planning'),
        ('retrospective', 'Retrospectives'),
        ('standup', 'Daily Standups'),
        ('estimation', 'Estimation'),
        ('story points', 'Story Points'),
        ('velocity', 'Velocity Tracking'),
        ('burndown', 'Burndown Charts'),
        ('release management', 'Release Management'),
        ('change management', 'Change Management'),
        ('stakeholder management', 'Stakeholder Management'),
        ('product management', 'Product Management'),
        ('project management', 'Project Management'),
        ('pmp', 'PMP'),
        ('prince2', 'PRINCE2'),
        ('safe', 'SAFe'),
        ('leads', 'LeSS'),
        ('scrum master', 'Scrum Master'),
        ('product owner', 'Product Owner'),
    ],
    'Mobile Development': [
        ('mobile development', 'Mobile Development'),
        ('ios', 'iOS Development'),
        ('android', 'Android Development'),
        ('swift', 'Swift'),
        ('objective-c', 'Objective-C'),
        ('kotlin', 'Kotlin'),
        ('java', 'Java (Android)'),
        ('flutter', 'Flutter'),
        ('dart', 'Dart'),
        ('react native', 'React Native'),
        ('expo', 'Expo'),
        ('xamarin', 'Xamarin'),
        ('maui', '.NET MAUI'),
        ('ionic', 'Ionic'),
        ('capacitor', 'Capacitor'),
        ('cordova', 'Cordova'),
        ('phonegap', 'PhoneGap'),
        ('native script', 'NativeScript'),
        ('app store', 'App Store'),
        ('play store', 'Google Play Store'),
        ('testflight', 'TestFlight'),
        ('firebase', 'Firebase'),
        ('push notifications', 'Push Notifications'),
        ('in-app purchases', 'In-App Purchases'),
        ('mobile ui', 'Mobile UI/UX'),
        ('material design', 'Material Design'),
        ('human interface guidelines', 'Human Interface Guidelines'),
        ('swiftui', 'SwiftUI'),
        ('uikit', 'UIKit'),
        ('jetpack compose', 'Jetpack Compose'),
        ('xml layout', 'XML Layouts'),
        ('gradle', 'Gradle (Android)'),
        ('cocoapods', 'CocoaPods'),
        ('swift package manager', 'Swift Package Manager'),
        ('fastlane', 'Fastlane'),
        ('bitrise', 'Bitrise'),
        ('codemagic', 'Codemagic'),
        ('app center', 'App Center'),
    ],
    'Cybersecurity': [
        ('cybersecurity', 'Cybersecurity'),
        ('security', 'Information Security'),
        ('infosec', 'Information Security'),
        ('penetration testing', 'Penetration Testing'),
        ('pentest', 'Penetration Testing'),
        ('ethical hacking', 'Ethical Hacking'),
        ('vulnerability assessment', 'Vulnerability Assessment'),
        ('security audit', 'Security Auditing'),
        ('owasp', 'OWASP'),
        ('owasp top 10', 'OWASP Top 10'),
        ('xss', 'XSS'),
        ('sql injection', 'SQL Injection'),
        ('csrf', 'CSRF'),
        ('authentication', 'Authentication Security'),
        ('authorization', 'Authorization'),
        ('oauth', 'OAuth Security'),
        ('jwt security', 'JWT Security'),
        ('encryption', 'Encryption'),
        ('cryptography', 'Cryptography'),
        ('tls', 'TLS/SSL'),
        ('certificate', 'Certificate Management'),
        ('pki', 'PKI'),
        ('hashing', 'Hashing'),
        ('bcrypt', 'bcrypt'),
        ('argon2', 'Argon2'),
        ('scrypt', 'scrypt'),
        ('siem', 'SIEM'),
        ('soc', 'SOC'),
        ('incident response', 'Incident Response'),
        ('threat modeling', 'Threat Modeling'),
        ('secure coding', 'Secure Coding'),
        ('devsecops', 'DevSecOps'),
        ('sast', 'SAST'),
        ('dast', 'DAST'),
        ('iast', 'IAST'),
        ('sca', 'SCA'),
        ('dependency scanning', 'Dependency Scanning'),
        ('container security', 'Container Security'),
        ('kubernetes security', 'Kubernetes Security'),
        ('cloud security', 'Cloud Security'),
        ('iam', 'IAM'),
        ('rbac', 'RBAC'),
        ('abac', 'ABAC'),
        ('zero trust', 'Zero Trust'),
        ('mfa', 'MFA'),
        ('2fa', '2FA'),
        ('sso', 'SSO'),
        ('saml', 'SAML'),
        ('openid connect', 'OpenID Connect'),
        ('ldap', 'LDAP'),
        ('active directory', 'Active Directory'),
    ],
    'Data Engineering': [
        ('data engineering', 'Data Engineering'),
        ('etl', 'ETL'),
        ('elt', 'ELT'),
        ('data pipeline', 'Data Pipelines'),
        ('data warehouse', 'Data Warehousing'),
        ('data lake', 'Data Lake'),
        ('lakehouse', 'Lakehouse'),
        ('apache spark', 'Apache Spark'),
        ('pyspark', 'PySpark'),
        ('spark sql', 'Spark SQL'),
        ('databricks', 'Databricks'),
        ('apache airflow', 'Apache Airflow'),
        ('airflow', 'Apache Airflow'),
        ('prefect', 'Prefect'),
        ('dagster', 'Dagster'),
        ('luigi', 'Luigi'),
        ('kafka', 'Apache Kafka'),
        ('kafka streams', 'Kafka Streams'),
        ('ksql', 'ksqlDB'),
        ('flink', 'Apache Flink'),
        ('beam', 'Apache Beam'),
        ('pulsar', 'Apache Pulsar'),
        ('rabbitmq', 'RabbitMQ'),
        ('aws kinesis', 'AWS Kinesis'),
        ('google pubsub', 'Google Pub/Sub'),
        ('azure event hubs', 'Azure Event Hubs'),
        ('cdc', 'Change Data Capture'),
        ('debezium', 'Debezium'),
        ('maxwell', 'Maxwell'),
        ('dbt', 'dbt'),
        ('data build tool', 'dbt'),
        ('sqlmesh', 'SQLMesh'),
        ('data quality', 'Data Quality'),
        ('great expectations', 'Great Expectations'),
        ('soda', 'Soda'),
        ('data contracts', 'Data Contracts'),
        ('data governance', 'Data Governance'),
        ('data catalog', 'Data Catalog'),
        ('data lineage', 'Data Lineage'),
        ('amundsen', 'Amundsen'),
        ('datahub', 'DataHub'),
        ('atlan', 'Atlan'),
        ('collibra', 'Collibra'),
        ('alation', 'Alation'),
        ('feature store', 'Feature Store'),
        ('feast', 'Feast'),
        ('hopsworks', 'Hopsworks'),
        ('tecton', 'Tecton'),
        ('mlflow', 'MLflow'),
        ('dvc', 'DVC'),
        ('delta lake', 'Delta Lake'),
        ('iceberg', 'Apache Iceberg'),
        ('hudi', 'Apache Hudi'),
        ('parquet', 'Parquet'),
        ('avro', 'Avro'),
        ('orc', 'ORC'),
        ('protobuf', 'Protocol Buffers'),
    ],
    'API Development & Integration': [
        ('api development', 'API Development'),
        ('rest api', 'REST APIs'),
        ('restful', 'RESTful APIs'),
        ('graphql', 'GraphQL'),
        ('grpc', 'gRPC'),
        ('websocket', 'WebSockets'),
        ('webhooks', 'Webhooks'),
        ('api gateway', 'API Gateway'),
        ('kong', 'Kong'),
        ('apigee', 'Apigee'),
        ('aws api gateway', 'AWS API Gateway'),
        ('azure api management', 'Azure API Management'),
        ('google cloud endpoints', 'Google Cloud Endpoints'),
        ('rate limiting', 'Rate Limiting'),
        ('throttling', 'Throttling'),
        ('api versioning', 'API Versioning'),
        ('api documentation', 'API Documentation'),
        ('swagger', 'Swagger/OpenAPI'),
        ('openapi', 'OpenAPI'),
        ('redoc', 'ReDoc'),
        ('postman', 'Postman'),
        ('insomnia', 'Insomnia'),
        ('hoppscotch', 'Hoppscotch'),
        ('api testing', 'API Testing'),
        ('contract testing', 'Contract Testing'),
        ('pact', 'Pact'),
        ('schema registry', 'Schema Registry'),
        ('avro', 'Avro'),
        ('protobuf', 'Protocol Buffers'),
        ('oauth', 'OAuth'),
        ('oauth2', 'OAuth2'),
        ('openid connect', 'OpenID Connect'),
        ('jwt', 'JWT'),
        ('api key', 'API Keys'),
        ('mutual tls', 'mTLS'),
        ('integration', 'System Integration'),
        ('microservices', 'Microservices'),
        ('service mesh', 'Service Mesh'),
        ('event driven', 'Event-Driven Architecture'),
        ('message queue', 'Message Queues'),
        ('event streaming', 'Event Streaming'),
        ('cqrs', 'CQRS'),
        ('event sourcing', 'Event Sourcing'),
    ],
    'UI/UX Design': [
        ('ui design', 'UI Design'),
        ('ux design', 'UX Design'),
        ('user interface', 'User Interface Design'),
        ('user experience', 'User Experience Design'),
        ('figma', 'Figma'),
        ('sketch', 'Sketch'),
        ('adobe xd', 'Adobe XD'),
        ('framer', 'Framer'),
        ('invision', 'InVision'),
        ('zeplin', 'Zeplin'),
        ('abstract', 'Abstract'),
        ('prototyping', 'Prototyping'),
        ('wireframing', 'Wireframing'),
        ('user research', 'User Research'),
        ('usability testing', 'Usability Testing'),
        ('user testing', 'User Testing'),
        ('a/b testing', 'A/B Testing'),
        ('design system', 'Design Systems'),
        ('design tokens', 'Design Tokens'),
        ('accessibility', 'Accessibility'),
        ('a11y', 'Accessibility'),
        ('wcag', 'WCAG'),
        ('responsive design', 'Responsive Design'),
        ('mobile first', 'Mobile-First Design'),
        ('design thinking', 'Design Thinking'),
        ('journey mapping', 'Journey Mapping'),
        ('personas', 'Personas'),
        ('information architecture', 'Information Architecture'),
        ('interaction design', 'Interaction Design'),
        ('visual design', 'Visual Design'),
        ('typography', 'Typography'),
        ('color theory', 'Color Theory'),
        ('iconography', 'Iconography'),
        ('illustration', 'Illustration'),
        ('motion design', 'Motion Design'),
        ('animation', 'Animation'),
        ('microinteractions', 'Microinteractions'),
        ('lottie', 'Lottie'),
        ('after effects', 'After Effects'),
        ('principle', 'Principle'),
        ('proto.io', 'Proto.io'),
        ('marvel', 'Marvel App'),
        ('balsamiq', 'Balsamiq'),
        ('mural', 'Mural'),
        ('miro', 'Miro'),
    ],
    'Game Development': [
        ('game development', 'Game Development'),
        ('unity', 'Unity'),
        ('unreal engine', 'Unreal Engine'),
        ('ue4', 'Unreal Engine 4'),
        ('ue5', 'Unreal Engine 5'),
        ('godot', 'Godot'),
        ('cocos2d', 'Cocos2d'),
        ('phaser', 'Phaser'),
        ('pixi.js', 'PixiJS'),
        ('three.js', 'Three.js'),
        ('babylon.js', 'Babylon.js'),
        ('c#', 'C# (Unity)'),
        ('c++', 'C++ (Unreal)'),
        ('blueprints', 'Blueprints (Unreal)'),
        ('gdscript', 'GDScript'),
        ('game design', 'Game Design'),
        ('level design', 'Level Design'),
        ('gameplay programming', 'Gameplay Programming'),
        ('graphics programming', 'Graphics Programming'),
        ('shader', 'Shader Programming'),
        ('hlsl', 'HLSL'),
        ('glsl', 'GLSL'),
        ('rendering', 'Rendering'),
        ('physics engine', 'Physics Engines'),
        ('havok', 'Havok'),
        ('physx', 'PhysX'),
        ('bullet physics', 'Bullet Physics'),
        ('ai in games', 'Game AI'),
        ('behavior trees', 'Behavior Trees'),
        ('state machines', 'State Machines'),
        ('navigation mesh', 'NavMesh'),
        ('multiplayer', 'Multiplayer Networking'),
        ('photon', 'Photon'),
        ('mirror', 'Mirror'),
        ('netcode', 'Netcode'),
        ('rollback', 'Rollback Netcode'),
        ('dedicated server', 'Dedicated Servers'),
        ('matchmaking', 'Matchmaking'),
        ('leaderboard', 'Leaderboards'),
        ('game analytics', 'Game Analytics'),
        ('unity analytics', 'Unity Analytics'),
        ('game monetization', 'Game Monetization'),
        ('in-app purchase', 'In-App Purchases'),
        ('ads mediation', 'Ads Mediation'),
        ('app store', 'App Store Connect'),
        ('play console', 'Google Play Console'),
        ('steamworks', 'Steamworks'),
        ('itch.io', 'itch.io'),
        ('game jam', 'Game Jams'),
    ],
    'Blockchain & Web3': [
        ('blockchain', 'Blockchain'),
        ('web3', 'Web3'),
        ('ethereum', 'Ethereum'),
        ('bitcoin', 'Bitcoin'),
        ('solidity', 'Solidity'),
        ('vyper', 'Vyper'),
        ('rust', 'Rust (Solana)'),
        ('solana', 'Solana'),
        ('anchor', 'Anchor'),
        ('cosmwasm', 'CosmWasm'),
        ('polkadot', 'Polkadot'),
        ('substrate', 'Substrate'),
        ('ink', 'ink!'),
        ('near', 'NEAR'),
        ('avalanche', 'Avalanche'),
        ('polygon', 'Polygon'),
        ('arbitrum', 'Arbitrum'),
        ('optimism', 'Optimism'),
        ('base', 'Base'),
        ('zksync', 'zkSync'),
        ('starknet', 'StarkNet'),
        ('smart contracts', 'Smart Contracts'),
        ('defi', 'DeFi'),
        ('nft', 'NFTs'),
        ('erc20', 'ERC-20'),
        ('erc721', 'ERC-721'),
        ('erc1155', 'ERC-1155'),
        ('dao', 'DAOs'),
        ('governance', 'Governance'),
        ('tokenomics', 'Tokenomics'),
        ('web3.js', 'web3.js'),
        ('ethers.js', 'ethers.js'),
        ('viem', 'viem'),
        ('wagmi', 'wagmi'),
        ('rainbowkit', 'RainbowKit'),
        ('walletconnect', 'WalletConnect'),
        ('metamask', 'MetaMask'),
        ('hardhat', 'Hardhat'),
        ('foundry', 'Foundry'),
        ('truffle', 'Truffle'),
        ('brownie', 'Brownie'),
        ('slither', 'Slither'),
        ('mythril', 'Mythril'),
        ('echidna', 'Echidna'),
        ('fuzzing', 'Smart Contract Fuzzing'),
        ('formal verification', 'Formal Verification'),
        ('zero knowledge', 'Zero Knowledge Proofs'),
        ('zk snark', 'zk-SNARKs'),
        ('zk stark', 'zk-STARKs'),
        ('plonk', 'PLONK'),
        ('circuit', 'Circuit Programming'),
        ('circom', 'Circom'),
        ('noir', 'Noir'),
        ('risc zero', 'RISC Zero'),
        ('sp1', 'SP1'),
    ],
    'Embedded Systems & IoT': [
        ('embedded systems', 'Embedded Systems'),
        ('iot', 'Internet of Things'),
        ('internet of things', 'IoT'),
        ('microcontroller', 'Microcontrollers'),
        ('mcu', 'MCU'),
        ('arm', 'ARM'),
        ('cortex-m', 'Cortex-M'),
        ('stm32', 'STM32'),
        ('esp32', 'ESP32'),
        ('esp8266', 'ESP8266'),
        ('arduino', 'Arduino'),
        ('raspberry pi', 'Raspberry Pi'),
        ('pico', 'Raspberry Pi Pico'),
        ('rp2040', 'RP2040'),
        ('nrf52', 'nRF52'),
        ('zephyr', 'Zephyr RTOS'),
        ('freertos', 'FreeRTOS'),
        ('threadx', 'ThreadX'),
        ('rtos', 'RTOS'),
        ('bare metal', 'Bare Metal Programming'),
        ('c', 'C (Embedded)'),
        ('c++', 'C++ (Embedded)'),
        ('rust', 'Rust (Embedded)'),
        ('assembly', 'Assembly'),
        ('gpio', 'GPIO'),
        ('i2c', 'I2C'),
        ('spi', 'SPI'),
        ('uart', 'UART'),
        ('can', 'CAN Bus'),
        ('usb', 'USB'),
        ('ethernet', 'Ethernet'),
        ('wifi', 'WiFi'),
        ('bluetooth', 'Bluetooth'),
        ('ble', 'Bluetooth Low Energy'),
        ('zigbee', 'Zigbee'),
        ('z-wave', 'Z-Wave'),
        ('lora', 'LoRa'),
        ('lorawan', 'LoRaWAN'),
        ('mqtt', 'MQTT'),
        ('coap', 'CoAP'),
        ('http', 'HTTP'),
        ('websocket', 'WebSocket'),
        ('sensor', 'Sensors'),
        ('actuator', 'Actuators'),
        ('adc', 'ADC'),
        ('dac', 'DAC'),
        ('pwm', 'PWM'),
        ('interrupt', 'Interrupts'),
        ('dma', 'DMA'),
        ('rtc', 'RTC'),
        ('watchdog', 'Watchdog Timer'),
        ('bootloader', 'Bootloaders'),
        ('dfu', 'DFU'),
        ('ota', 'OTA Updates'),
        ('firmware', 'Firmware Development'),
        ('device driver', 'Device Drivers'),
        ('kernel', 'Kernel Development'),
        ('linux kernel', 'Linux Kernel'),
        ('yocto', 'Yocto Project'),
        ('buildroot', 'Buildroot'),
        ('openwrt', 'OpenWrt'),
        ('pcb design', 'PCB Design'),
        ('kicad', 'KiCad'),
        ('eagle', 'EAGLE'),
        ('altium', 'Altium Designer'),
        ('easyeda', 'EasyEDA'),
        ('schematic', 'Schematic Capture'),
        ('layout', 'PCB Layout'),
    ],
}


def _seed_default_content():
    content = LandingPageContent.objects.first()
    if not content:
        content = LandingPageContent.objects.create()

    if not FeatureItem.objects.exists():
        FeatureItem.objects.bulk_create([
            FeatureItem(title='ATS optimized', description='Generate a clean CV that is easy for recruiting systems to read.', icon='⚡', order=1),
            FeatureItem(title='Job-targeted content', description='Shape your summary and skill list around the exact role you want.', icon='🎯', order=2),
            FeatureItem(title='Instant downloads', description='Export your final CV as PDF or DOCX in one click.', icon='📄', order=3),
        ])

    if not ProcessStep.objects.exists():
        ProcessStep.objects.bulk_create([
            ProcessStep(title='Share your details', description='Enter your profile, background, and target role.', order=1),
            ProcessStep(title='Match the role', description='Add job requirements and skills to align the content.', order=2),
            ProcessStep(title='Generate and download', description='Preview the ATS-ready result and export it instantly.', order=3),
        ])

    if not BenefitItem.objects.exists():
        BenefitItem.objects.bulk_create([
            BenefitItem(title='Keyword optimization', description='Use the job description to include relevant terms naturally.', icon='🔑', order=1),
            BenefitItem(title='Professional formatting', description='Create a structured CV with strong sections and easy reading.', icon='🧩', order=2),
            BenefitItem(title='No account required', description='No signup, no login, and no long-term data storage.', icon='🔒', order=3),
        ])

    if not Testimonial.objects.exists():
        Testimonial.objects.bulk_create([
            Testimonial(quote='The output looked polished and recruiter-friendly from the very first draft.', author='Aisha K.', role='Software Engineer', order=1),
            Testimonial(quote='I used it to tailor my CV for three roles in one afternoon.', author='David J.', role='Marketing Lead', order=2),
        ])

    if not FAQItem.objects.exists():
        FAQItem.objects.bulk_create([
            FAQItem(question='Is my data safe?', answer='Yes. We do not store your information beyond the current session and you can download your CV immediately.', order=1),
            FAQItem(question='Can I export to PDF or DOCX?', answer='Yes. The preview page includes both options for quick downloads.', order=2),
        ])

    return content


def _get_builder_data(request):
    data = request.session.get('builder_data')
    if not data:
        data = deepcopy(EMPTY_BUILDER_DATA)
        request.session['builder_data'] = data
    return data


def _extract_skill_categories(job_description='', required_skills='', extra_skills='', keywords=''):
    text = ' '.join([job_description, required_skills, extra_skills, keywords]).lower()
    text = re.sub(r'[^a-z0-9+\s./-]', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()

    categories = []
    for category in DEFAULT_SKILL_CATEGORIES:
        matches = []
        seen = set()
        for raw_term, display_term in SKILL_KEYWORD_MAP.get(category, []):
            if not raw_term:
                continue
            if raw_term in text:
                if display_term not in seen:
                    seen.add(display_term)
                    matches.append(display_term)

        if matches:
            categories.append({'category': category, 'items': ', '.join(matches)})

    if not categories:
        categories = [
            {'category': 'Project Management & Soft Skills', 'items': 'Communication, Problem Solving'},
        ]

    return categories


def _generate_summary(data):
    title = data.get('job_requirements', {}).get('job_title') or data.get('job_requirements', {}).get('role') or ''
    job_description = data.get('job_requirements', {}).get('job_description', '')
    skills = []
    for category in data.get('skill_categories', []):
        items = [item.strip() for item in category.get('items', '').split(',') if item.strip()]
        skills.extend(items)
    top_skills = ', '.join(skills[:6]) if skills else ''
    
    # If user provided a custom summary in job requirements, use that
    custom_summary = data.get('job_requirements', {}).get('custom_summary', '').strip()
    if custom_summary:
        return custom_summary
    
    # Generate summary based on job description and skills
    if title and top_skills:
        return f"{title} professional with expertise in {top_skills}. Proven track record of delivering high-impact solutions through technical excellence and collaborative problem-solving. Passionate about leveraging technology to drive business value and innovation."
    elif title:
        return f"{title} professional with a strong foundation in relevant technologies and methodologies. Experienced in delivering quality results through analytical thinking and effective collaboration. Committed to continuous learning and professional growth."
    elif top_skills:
        return f"Skilled professional with expertise in {top_skills}. Demonstrated ability to solve complex problems and deliver results in fast-paced environments. Strong communicator and team player committed to excellence."
    return 'Motivated professional seeking to leverage technical skills and experience to contribute to organizational success. Adaptable learner with strong problem-solving abilities and a passion for continuous improvement.'


def _calculate_score(data):
    keywords = [item.strip().lower() for item in data['job_requirements'].get('keywords', '').split(',') if item.strip()]
    required_skills = [item.strip().lower() for item in data['job_requirements'].get('required_skills', '').split(',') if item.strip()]
    all_keywords = keywords + required_skills

    text = ' '.join([
        data['job_requirements'].get('job_title', ''),
        data['job_requirements'].get('job_description', ''),
        data['job_requirements'].get('role', ''),
        ' '.join([item.get('items', '') for item in data.get('skill_categories', [])]),
        ' '.join([project.get('title', '') + ' ' + project.get('description', '') for project in data.get('projects', [])]),
        ' '.join([edu.get('institution', '') + ' ' + edu.get('degree', '') + ' ' + edu.get('details', '') for edu in data.get('education', [])]),
        ' '.join([exp.get('company', '') + ' ' + exp.get('role', '') + ' ' + exp.get('responsibilities', '') for exp in data.get('experience', [])]),
    ]).lower()

    matched_keywords = []
    for keyword in all_keywords:
        if keyword and re.search(r'\b' + re.escape(keyword) + r'\b', text):
            matched_keywords.append(keyword)

    keyword_total = max(1, len(all_keywords))
    keyword_count = len(matched_keywords)
    score = 55 + round((keyword_count / keyword_total) * 30)
    if data['personal'].get('full_name') and data['personal'].get('email') and data['personal'].get('phone'):
        score += 8
    if data['job_requirements'].get('job_title') and data['job_requirements'].get('job_description'):
        score += 7
    if data.get('skill_categories'):
        score += 5
    if data.get('projects') or data.get('experience') or data.get('education'):
        score += 5
    score = min(score, 98)
    suggestions = [kw for kw in all_keywords if kw and not re.search(r'\b' + re.escape(kw) + r'\b', text)]
    return score, suggestions, keyword_count, keyword_total


def landing_page(request):
    content = _seed_default_content()
    context = {
        'content': content,
        'features': FeatureItem.objects.all(),
        'steps': ProcessStep.objects.all(),
        'benefits': BenefitItem.objects.all(),
        'testimonials': Testimonial.objects.all(),
        'faqs': FAQItem.objects.all(),
    }
    return render(request, 'cv_app/landing.html', context)


def builder(request):
    step = int(request.GET.get('step', 1))
    data = _get_builder_data(request)

    if request.method == 'POST':
        step = int(request.POST.get('step', step))

        if step == 1:
            data['personal'] = {
                'full_name': request.POST.get('full_name', '').strip(),
                'email': request.POST.get('email', '').strip(),
                'phone': request.POST.get('phone', '').strip(),
                'linkedin': request.POST.get('linkedin', '').strip(),
                'github': request.POST.get('github', '').strip(),
                'portfolio': request.POST.get('portfolio', '').strip(),
                'location': request.POST.get('location', '').strip(),
            }

        elif step == 2:
            data['job_requirements'] = {
                'job_title': request.POST.get('job_title', '').strip(),
                'role': request.POST.get('role', '').strip(),
                'job_description': request.POST.get('job_description', '').strip(),
                'keywords': request.POST.get('keywords', '').strip(),
                'required_skills': request.POST.get('required_skills', '').strip(),
                'custom_summary': request.POST.get('custom_summary', '').strip(),
            }
            data['skill_categories'] = _extract_skill_categories(
                data['job_requirements'].get('job_description', ''),
                data['job_requirements'].get('required_skills', ''),
                data['job_requirements'].get('required_skills', ''),
                data['job_requirements'].get('keywords', ''),
            )

        elif step == 3:
            skill_categories = []
            category_names = request.POST.getlist('skill_category[]')
            category_items = request.POST.getlist('skill_items[]')
            custom_categories = request.POST.getlist('skill_custom_category[]')
            for i, (cat, items) in enumerate(zip(category_names, category_items)):
                # If "Other" is selected, use the custom category
                if cat.strip() == 'Other' and i < len(custom_categories) and custom_categories[i].strip():
                    cat = custom_categories[i].strip()
                if cat.strip() or items.strip():
                    skill_categories.append({'category': cat.strip(), 'items': items.strip()})
            data['skill_categories'] = skill_categories if skill_categories else data.get('skill_categories', [])

        elif step == 4:
            project_titles = request.POST.getlist('project_title[]')
            project_categories = request.POST.getlist('project_category[]')
            project_descriptions = request.POST.getlist('project_description[]')
            source_code_links = request.POST.getlist('source_code_link[]')
            project_links = request.POST.getlist('project_link[]')

            data['projects'] = [
                {
                    'title': title.strip(),
                    'category': category.strip(),
                    'description': description.strip(),
                    'source_code_link': source_code.strip(),
                    'project_link': link.strip(),
                }
                for title, category, description, source_code, link in zip(
                    project_titles, project_categories, project_descriptions, source_code_links, project_links
                )
                if title.strip() or category.strip() or description.strip() or source_code.strip() or link.strip()
            ]

        elif step == 5:
            company_names = request.POST.getlist('experience_company[]')
            experience_roles = request.POST.getlist('experience_role[]')
            experience_periods = request.POST.getlist('experience_period[]')
            experience_responsibilities = request.POST.getlist('experience_responsibilities[]')

            if company_names or experience_roles:
                data['experience'] = [
                    {
                        'company': company.strip(),
                        'role': role.strip(),
                        'period': period.strip(),
                        'responsibilities': responsibilities.strip(),
                    }
                    for company, role, period, responsibilities in zip(
                        company_names, experience_roles, experience_periods, experience_responsibilities
                    )
                    if company.strip() or role.strip() or period.strip() or responsibilities.strip()
                ]

        elif step == 6:
            degrees = request.POST.getlist('education_degree[]')
            institutions = request.POST.getlist('education_institution[]')
            periods = request.POST.getlist('education_period[]')
            gpa_cgpas = request.POST.getlist('education_gpa_cgpa[]')
            gpa_maxes = request.POST.getlist('education_gpa_max[]')
            education_details = request.POST.getlist('education_details[]')

            data['education'] = [
                {
                    'degree': degree.strip(),
                    'institution': institution.strip(),
                    'period': period.strip(),
                    'gpa_cgpa': gpa_cgpa.strip(),
                    'gpa_max': gpa_max.strip(),
                    'details': details.strip(),
                }
                for degree, institution, period, gpa_cgpa, gpa_max, details in zip(
                    degrees, institutions, periods, gpa_cgpas, gpa_maxes, education_details
                )
                if degree.strip() or institution.strip() or period.strip() or gpa_cgpa.strip() or gpa_max.strip() or details.strip()
            ]

        elif step == 7:
            cert_names = request.POST.getlist('certification_name[]')
            cert_issuers = request.POST.getlist('certification_issuer[]')
            cert_dates = request.POST.getlist('certification_date[]')
            cert_expirations = request.POST.getlist('certification_expiration[]')
            cert_ids = request.POST.getlist('certification_credential_id[]')
            cert_urls = request.POST.getlist('certification_credential_url[]')

            data['certifications'] = [
                {
                    'name': name.strip(),
                    'issuer': issuer.strip(),
                    'date': date.strip(),
                    'expiration': expiration.strip(),
                    'credential_id': cred_id.strip(),
                    'credential_url': cred_url.strip(),
                }
                for name, issuer, date, expiration, cred_id, cred_url in zip(
                    cert_names, cert_issuers, cert_dates, cert_expirations, cert_ids, cert_urls
                )
                if name.strip() or issuer.strip() or date.strip()
            ]

        elif step == 8:
            honor_titles = request.POST.getlist('honor_title[]')
            honor_issuers = request.POST.getlist('honor_issuer[]')
            honor_dates = request.POST.getlist('honor_date[]')
            honor_descriptions = request.POST.getlist('honor_description[]')

            data['honors_awards'] = [
                {
                    'title': title.strip(),
                    'issuer': issuer.strip(),
                    'date': date.strip(),
                    'description': description.strip(),
                }
                for title, issuer, date, description in zip(
                    honor_titles, honor_issuers, honor_dates, honor_descriptions
                )
                if title.strip() or issuer.strip() or date.strip()
            ]

        request.session['builder_data'] = data

        if step < 9:
            return redirect(f'/create/?step={step + 1}')
        return redirect('/preview/')

    if not data.get('skill_categories') and data.get('job_requirements', {}).get('job_description'):
        data['skill_categories'] = _extract_skill_categories(
            data['job_requirements'].get('job_description', ''),
            data['job_requirements'].get('required_skills', ''),
            data['job_requirements'].get('required_skills', ''),
            data['job_requirements'].get('keywords', ''),
        )
        request.session['builder_data'] = data

    context = {
        'step': step,
        'data': data,
        'skill_categories': data.get('skill_categories', []),
        'default_skill_categories': DEFAULT_SKILL_CATEGORIES,
    }
    return render(request, 'cv_app/builder.html', context)


def preview(request):
    data = _get_builder_data(request)
    score, suggestions, keyword_count, keyword_total = _calculate_score(data)
    context = {
        'data': data,
        'score': score,
        'keyword_count': keyword_count,
        'keyword_total': keyword_total,
        'keyword_suggestions': suggestions[:6],
        'generated_summary': _generate_summary(data),
    }
    return render(request, 'cv_app/preview.html', context)


def download_pdf(request):
    data = _get_builder_data(request)
    generated_summary = _generate_summary(data)

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=0.75 * inch,
        leftMargin=0.75 * inch,
        topMargin=0.75 * inch,
        bottomMargin=0.75 * inch,
    )
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=20, leading=24, textColor=colors.HexColor('#1d4ed8'))
    heading_style = ParagraphStyle('Heading', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=13, leading=15, textColor=colors.HexColor('#111827'))
    subheading_style = ParagraphStyle('SubHeading', parent=styles['Heading3'], fontName='Helvetica-Bold', fontSize=11, leading=13, textColor=colors.HexColor('#1d4ed8'))
    body_style = ParagraphStyle('Body', parent=styles['BodyText'], fontName='Helvetica', fontSize=10, leading=13, textColor=colors.HexColor('#111827'))
    small_style = ParagraphStyle('Small', parent=styles['BodyText'], fontName='Helvetica', fontSize=9, leading=11, textColor=colors.HexColor('#4b5563'))
    bullet_style = ParagraphStyle('Bullet', parent=body_style, leftIndent=20, bulletIndent=10)

    story = []

    name = data['personal'].get('full_name') or 'Your Name'
    story.append(Paragraph(name, title_style))

    contact_parts = []
    if data['personal'].get('location'):
        contact_parts.append(data['personal'].get('location'))
    if data['personal'].get('email'):
        contact_parts.append(data['personal'].get('email'))
    if data['personal'].get('phone'):
        contact_parts.append(data['personal'].get('phone'))
    story.append(Paragraph(' | '.join(contact_parts), small_style))

    links = []
    if data['personal'].get('linkedin'):
        links.append(f"LinkedIn: {data['personal'].get('linkedin')}")
    if data['personal'].get('github'):
        links.append(f"GitHub: {data['personal'].get('github')}")
    if data['personal'].get('portfolio'):
        links.append(f"Portfolio: {data['personal'].get('portfolio')}")
    if links:
        story.append(Paragraph(' • '.join(links), small_style))

    story.append(Spacer(1, 8))
    story.append(Paragraph('PROFESSIONAL SUMMARY', heading_style))
    story.append(Paragraph(generated_summary, body_style))
    story.append(Spacer(1, 8))

    story.append(Paragraph('PROFESSIONAL SKILLS', heading_style))
    for category in data.get('skill_categories', []):
        if category.get('items'):
            story.append(Paragraph(f"<b>{category.get('category')}</b>: {category.get('items')}", body_style))
    if not data.get('skill_categories'):
        story.append(Paragraph('<b>Communication</b>: Communication, Problem Solving', body_style))

    story.append(Spacer(1, 8))
    story.append(Paragraph('PROJECTS', heading_style))
    for project in data.get('projects', []):
        if not project.get('title') and not project.get('description'):
            continue

        title_parts = [f"<b>{project.get('title') or 'Project'}</b>"]
        if project.get('source_code_link'):
            title_parts.append(" <b>[Source Code]</b>")
        story.append(Paragraph(''.join(title_parts), body_style))

        if project.get('description'):
            bullets = []
            for line in project.get('description').splitlines():
                line = line.strip()
                if line:
                    bullets.append(line)
            if not bullets:
                bullets = [project.get('description').strip()]

            story.append(
                ListFlowable(
                    [ListItem(Paragraph(f"• {b}", bullet_style), bulletColor=colors.HexColor('#2563eb')) for b in bullets],
                    bulletType='bullet',
                )
            )

        links = []
        if project.get('source_code_link'):
            links.append(f"[Source Code] {project.get('source_code_link')}")
        if project.get('project_link'):
            links.append(f"Project Link: {project.get('project_link')}")
        if links:
            story.append(Paragraph(' • '.join(links), small_style))

        story.append(Spacer(1, 4))

    if not data.get('projects'):
        story.append(Paragraph('No project details entered', body_style))

    story.append(Spacer(1, 8))
    story.append(Paragraph('EDUCATION', heading_style))

    for education in data.get('education', []):
        if not education.get('institution') and not education.get('degree'):
            continue
        story.append(Paragraph(f"<b>{education.get('degree') or 'Degree'}</b> — {education.get('institution') or 'Institution'} | <b>Graduated</b>: {education.get('period')}", body_style))
        if education.get('gpa_cgpa') or education.get('gpa_max'):
            gpa_text = f"CGPA/GPA: {education.get('gpa_cgpa')}"
            if education.get('gpa_max'):
                gpa_text += f" / {education.get('gpa_max')}"
            story.append(Paragraph(gpa_text, small_style))
        if education.get('details'):
            story.append(Paragraph(education.get('details'), body_style))
        story.append(Spacer(1, 4))

    if not data.get('education'):
        story.append(Paragraph('No education details entered', body_style))

    story.append(Spacer(1, 8))
    story.append(Paragraph('EXPERIENCE', heading_style))
    for experience in data.get('experience', []):
        if not experience.get('company') and not experience.get('role'):
            continue
        story.append(Paragraph(f"<b>{experience.get('role') or 'Role'}</b> @ {experience.get('company') or 'Company'}", body_style))
        if experience.get('period'):
            story.append(Paragraph(experience.get('period'), small_style))
        responsibilities = [item.strip() for item in experience.get('responsibilities', '').split('\n') if item.strip()]
        if responsibilities:
            story.append(
                ListFlowable(
                    [ListItem(Paragraph(f"• {item}", bullet_style), bulletColor=colors.HexColor('#2563eb')) for item in responsibilities],
                    bulletType='bullet',
                )
            )
        story.append(Spacer(1, 4))

    if not data.get('experience'):
        story.append(Paragraph('No experience details entered', body_style))

    # Certifications
    if data.get('certifications'):
        story.append(Spacer(1, 8))
        story.append(Paragraph('CERTIFICATIONS', heading_style))
        for cert in data.get('certifications', []):
            if not cert.get('name') and not cert.get('issuer'):
                continue
            cert_parts = [f"<b>{cert.get('name') or 'Certification'}</b> — {cert.get('issuer') or 'Issuer'}"]
            if cert.get('date'):
                cert_parts.append(f" | <b>Issued</b>: {cert.get('date')}")
            if cert.get('expiration'):
                cert_parts.append(f" | <b>Expires</b>: {cert.get('expiration')}")
            if cert.get('credential_id'):
                cert_parts.append(f" | <b>Credential ID</b>: {cert.get('credential_id')}")
            story.append(Paragraph(''.join(cert_parts), body_style))
            if cert.get('credential_url'):
                story.append(Paragraph(f"Verify: {cert.get('credential_url')}", small_style))
            story.append(Spacer(1, 4))

    # Honors & Awards
    if data.get('honors_awards'):
        story.append(Spacer(1, 8))
        story.append(Paragraph('HONORS & AWARDS', heading_style))
        for honor in data.get('honors_awards', []):
            if not honor.get('title') and not honor.get('issuer'):
                continue
            honor_parts = [f"<b>{honor.get('title') or 'Award'}</b> — {honor.get('issuer') or 'Issuer'}"]
            if honor.get('date'):
                honor_parts.append(f" | <b>Date</b>: {honor.get('date')}")
            story.append(Paragraph(''.join(honor_parts), body_style))
            if honor.get('description'):
                story.append(Paragraph(honor.get('description'), body_style))
            story.append(Spacer(1, 4))

    doc.build(story)
    pdf_data = buffer.getvalue()
    response = HttpResponse(pdf_data, content_type='application/pdf')
    response['Content-Disposition'] = 'attachment; filename="cv_guide_cv.pdf"'
    return response


def download_docx(request):
    data = _get_builder_data(request)
    document = Document()

    style = document.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)
    font.color.rgb = RGBColor(0x14, 0x21, 0x3D)

    # Name
    heading = document.add_heading(data['personal'].get('full_name') or 'Your Name', level=0)
    heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in heading.runs:
        run.font.color.rgb = RGBColor(0x1D, 0x4E, 0xD8)
        run.font.size = Pt(24)

    # Contact info
    contact_parts = []
    if data['personal'].get('location'):
        contact_parts.append(data['personal'].get('location'))
    if data['personal'].get('email'):
        contact_parts.append(data['personal'].get('email'))
    if data['personal'].get('phone'):
        contact_parts.append(data['personal'].get('phone'))
    p = document.add_paragraph(' | '.join(contact_parts))
    p.style.font.size = Pt(10)
    p.style.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)

    links = []
    if data['personal'].get('linkedin'):
        links.append(f"LinkedIn: {data['personal'].get('linkedin')}")
    if data['personal'].get('github'):
        links.append(f"GitHub: {data['personal'].get('github')}")
    if data['personal'].get('portfolio'):
        links.append(f"Portfolio: {data['personal'].get('portfolio')}")
    if links:
        p = document.add_paragraph(' • '.join(links))
        p.style.font.size = Pt(10)
        p.style.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)

    document.add_paragraph()  # spacer

    # Professional Summary
    heading = document.add_heading('PROFESSIONAL SUMMARY', level=1)
    for run in heading.runs:
        run.font.color.rgb = RGBColor(0x1D, 0x4E, 0xD8)
        run.font.size = Pt(14)
    document.add_paragraph(_generate_summary(data))

    # Professional Skills
    heading = document.add_heading('PROFESSIONAL SKILLS', level=1)
    for run in heading.runs:
        run.font.color.rgb = RGBColor(0x1D, 0x4E, 0xD8)
        run.font.size = Pt(14)
    for category in data.get('skill_categories', []):
        if category.get('items'):
            p = document.add_paragraph()
            run = p.add_run(f"- {category.get('category')}: ")
            run.bold = True
            p.add_run(category.get('items'))

    # Projects
    heading = document.add_heading('PROJECTS', level=1)
    for run in heading.runs:
        run.font.color.rgb = RGBColor(0x1D, 0x4E, 0xD8)
        run.font.size = Pt(14)
    for project in data.get('projects', []):
        if project.get('title') or project.get('description'):
            p = document.add_paragraph()
            run = p.add_run(project.get('title') or 'Project')
            run.bold = True
            if project.get('source_code_link'):
                run2 = p.add_run(' [Source Code]')
                run2.bold = True

            if project.get('description'):
                lines = [ln.strip() for ln in project.get('description').splitlines() if ln.strip()]
                if not lines:
                    lines = [project.get('description').strip()]
                for b in lines:
                    document.add_paragraph(b, style='List Bullet')

            link_parts = []
            if project.get('source_code_link'):
                link_parts.append(f"[Source Code] {project.get('source_code_link')}")
            if project.get('project_link'):
                link_parts.append(f"Project Link: {project.get('project_link')}")
            if link_parts:
                p = document.add_paragraph(' • '.join(link_parts))
                p.style.font.size = Pt(9)
                p.style.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)

    # Education
    heading = document.add_heading('EDUCATION', level=1)
    for run in heading.runs:
        run.font.color.rgb = RGBColor(0x1D, 0x4E, 0xD8)
        run.font.size = Pt(14)
    for education in data.get('education', []):
        if education.get('institution') or education.get('degree'):
            degree = education.get('degree') or 'Degree'
            inst = education.get('institution') or 'Institution'
            year = education.get('period') or ''
            p = document.add_paragraph()
            run = p.add_run(f"{degree} | {inst} | Graduated: {year}")
            run.bold = True

            gpa = education.get('gpa_cgpa')
            gmax = education.get('gpa_max')
            if gpa or gmax:
                if gmax:
                    document.add_paragraph(f"CGPA/GPA: {gpa} / {gmax}")
                else:
                    document.add_paragraph(f"CGPA/GPA: {gpa}")

            if education.get('details'):
                document.add_paragraph(education.get('details'))

    # Experience
    heading = document.add_heading('EXPERIENCE', level=1)
    for run in heading.runs:
        run.font.color.rgb = RGBColor(0x1D, 0x4E, 0xD8)
        run.font.size = Pt(14)
    for experience in data.get('experience', []):
        if experience.get('company') or experience.get('role'):
            p = document.add_paragraph()
            run = p.add_run(f"{experience.get('role') or 'Role'} @ {experience.get('company') or 'Company'}")
            run.bold = True
            if experience.get('period'):
                document.add_paragraph(experience.get('period'))
            for responsibility in [item.strip() for item in experience.get('responsibilities', '').split('\n') if item.strip()]:
                document.add_paragraph(responsibility, style='List Bullet')

    # Certifications
    if data.get('certifications'):
        heading = document.add_heading('CERTIFICATIONS', level=1)
        for run in heading.runs:
            run.font.color.rgb = RGBColor(0x1D, 0x4E, 0xD8)
            run.font.size = Pt(14)
        for cert in data.get('certifications', []):
            if cert.get('name') or cert.get('issuer'):
                p = document.add_paragraph()
                run = p.add_run(f"{cert.get('name') or 'Certification'} — {cert.get('issuer') or 'Issuer'}")
                run.bold = True
                details = []
                if cert.get('date'):
                    details.append(f"Issued: {cert.get('date')}")
                if cert.get('expiration'):
                    details.append(f"Expires: {cert.get('expiration')}")
                if cert.get('credential_id'):
                    details.append(f"Credential ID: {cert.get('credential_id')}")
                if details:
                    document.add_paragraph(' | '.join(details))
                if cert.get('credential_url'):
                    p = document.add_paragraph(f"Verify: {cert.get('credential_url')}")
                    p.style.font.size = Pt(9)
                    p.style.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)

    # Honors & Awards
    if data.get('honors_awards'):
        heading = document.add_heading('HONORS & AWARDS', level=1)
        for run in heading.runs:
            run.font.color.rgb = RGBColor(0x1D, 0x4E, 0xD8)
            run.font.size = Pt(14)
        for honor in data.get('honors_awards', []):
            if honor.get('title') or honor.get('issuer'):
                p = document.add_paragraph()
                run = p.add_run(f"{honor.get('title') or 'Award'} — {honor.get('issuer') or 'Issuer'}")
                run.bold = True
                if honor.get('date'):
                    document.add_paragraph(honor.get('date'))
                if honor.get('description'):
                    document.add_paragraph(honor.get('description'))

    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    response['Content-Disposition'] = 'attachment; filename="cv_guide_cv.docx"'
    document.save(response)
    return response


@login_required(login_url='/cv-admin/login/')
def cv_admin_dashboard(request):
    content = LandingPageContent.objects.first() or LandingPageContent.objects.create()
    FeatureFormSet = modelformset_factory(FeatureItem, form=FeatureItemForm, extra=1)
    ProcessFormSet = modelformset_factory(ProcessStep, form=ProcessStepForm, extra=1)
    BenefitFormSet = modelformset_factory(BenefitItem, form=BenefitItemForm, extra=1)
    TestimonialFormSet = modelformset_factory(Testimonial, form=TestimonialForm, extra=1)
    FAQFormSet = modelformset_factory(FAQItem, form=FAQItemForm, extra=1)

    if request.method == 'POST':
        main_form = LandingPageContentForm(request.POST, instance=content)
        features_formset = FeatureFormSet(request.POST, queryset=FeatureItem.objects.all())
        steps_formset = ProcessFormSet(request.POST, queryset=ProcessStep.objects.all())
        benefits_formset = BenefitFormSet(request.POST, queryset=BenefitItem.objects.all())
        testimonials_formset = TestimonialFormSet(request.POST, queryset=Testimonial.objects.all())
        faqs_formset = FAQFormSet(request.POST, queryset=FAQItem.objects.all())
        if all([
            main_form.is_valid(),
            features_formset.is_valid(),
            steps_formset.is_valid(),
            benefits_formset.is_valid(),
            testimonials_formset.is_valid(),
            faqs_formset.is_valid(),
        ]):
            main_form.save()
            features_formset.save()
            steps_formset.save()
            benefits_formset.save()
            testimonials_formset.save()
            faqs_formset.save()
            messages.success(request, 'Landing page content updated.')
            return redirect('cv_admin_dashboard')
    else:
        main_form = LandingPageContentForm(instance=content)
        features_formset = FeatureFormSet(queryset=FeatureItem.objects.all())
        steps_formset = ProcessFormSet(queryset=ProcessStep.objects.all())
        benefits_formset = BenefitFormSet(queryset=BenefitItem.objects.all())
        testimonials_formset = TestimonialFormSet(queryset=Testimonial.objects.all())
        faqs_formset = FAQFormSet(queryset=FAQItem.objects.all())

    context = {
        'main_form': main_form,
        'features_formset': features_formset,
        'steps_formset': steps_formset,
        'benefits_formset': benefits_formset,
        'testimonials_formset': testimonials_formset,
        'faqs_formset': faqs_formset,
    }
    return render(request, 'cv_app/cv_admin.html', context)


def cv_admin_login(request):
    if request.user.is_authenticated:
        return redirect('cv_admin_dashboard')
    error = None
    if request.method == 'POST':
        username = request.POST.get('username', '').strip()
        password = request.POST.get('password', '').strip()
        user = authenticate(request, username=username, password=password)
        if user is not None and user.is_staff:
            login(request, user)
            return redirect('cv_admin_dashboard')
        error = 'Please use a valid staff account.'
    return render(request, 'cv_app/admin_login.html', {'error': error})


def cv_admin_logout(request):
    logout(request)
    return redirect('cv_admin_login')