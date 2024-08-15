pipeline {
    agent any

    environment {
        NODEJS_HOME = tool name: 'NodeJS'
        PATH = "${NODEJS_HOME}/bin:${env.PATH}"
    }

    stages {
        stage('Checkout') {
            steps {
                git branch: 'main', url: 'https://github.com/JAlikhani/Html2Docx'
            }
        }
        stage('Install Dependencies') {
            steps {
                sh 'npm install'
            }
        }
        stage('Run Tests') {
            steps {
                sh 'npm test'
            }
        }
        stage('Build') {
            steps {
                sh 'npm run build'
            }
        }
        stage('Deploy') {
            steps {
                sh '''
                # Copy the build files to the deployment directory
                sudo rm -rf /var/www/html/*
                sudo cp -r build/* /var/www/html/
                '''
            }
        }
    }

    post {
        always {
            cleanWs()
        }
    }
}