# GlobalSQA_Automation_Project
End-to-end test automation suite for the GlobalSQA website using UFT, covering data-driven and cross-browser testing, error handling, and modular test frameworks. Integrated with CI/CD pipelines through Jenkins for parallel execution.
### Disclaimer
This project was developed using a trial version of UFT One. It is intended for educational purposes only. Please ensure you have the proper licensing for UFT if you plan to use this code in a commercial setting.

## CI/CD Integration

This project integrates with Jenkins for continuous integration and deployment. Below are the key features of the CI/CD setup:

- **Jenkins Setup**: Automated execution of UFT test scripts is configured in Jenkins. Each commit triggers a new build that runs the UFT tests to ensure that no functionality is broken.
  
- **Build Triggers**: Tests are executed on every push to the `main` branch and on a scheduled basis to ensure regular testing.

- **Test Execution**: The UFT test scripts are run using the command line with the following command:
C:\Program Files (x86)\OpenText\UFT One\bin\UFTBatchRunner.exe -source G:\UFT Projects\GlobalSqa_Automation_Project\Scripts\TestScripts\DTS_LoginTest\Test.tsp


- **Results Reporting**: Test results are captured and reported back in Jenkins, allowing for easy tracking of test status.

### Screenshots


