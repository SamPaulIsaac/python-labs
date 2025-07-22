import os
import pandas as pd

# Prepare the data for the Excel file
topics_data = [
    ("Spring Boot Core Concepts", [
        "Auto-Configuration: In-depth understanding of how Spring Boot auto-configures beans and how to disable/override it for custom configurations.",
        "Starters: Leveraging Spring Boot starters and creating custom starters for reusable configurations and dependencies.",
        "CLI (Command-Line Interface): Using the Spring Boot CLI for quick prototyping and automating tasks.",
        "Annotations & Related Concepts: Advanced usage of @SpringBootApplication, @ConfigurationProperties, @Value, and understanding the lifecycle of Spring Boot applications."
    ]),
    ("Spring Boot Configuration", [
        "External Configuration Management: Managing properties and configuration files, Profiles for different environments.",
        "Custom Configuration Classes: Creating custom beans, using @ConfigurationProperties.",
        "Spring Cloud Config: Using for centralized configuration in distributed systems."
    ]),
    ("Spring Boot Web Development", [
        "Developing REST APIs with Spring MVC.",
        "Exception Handling using @ControllerAdvice.",
        "Full-Stack Development with Thymeleaf, WebFlux, WebSocket.",
        "Advanced Web Concepts: HATEOAS, CORS, file uploads, large payloads."
    ]),
    ("Spring Boot Data & Persistence", [
        "RDBMS with JPA, query tuning, lazy loading.",
        "NoSQL: MongoDB, Cassandra, Couchbase, Redis.",
        "Spring Data custom repositories.",
        "Transactional Management: propagation, isolation, JTA."
    ]),
    ("Spring Boot Security", [
        "Authentication, authorization, CSRF protection.",
        "JWT & OAuth2 integration.",
        "Method-Level Security with @PreAuthorize, @Secured.",
        "Custom Authentication filters."
    ]),
    ("Spring Boot Microservices", [
        "Spring Cloud: Eureka, Config, Gateway.",
        "Resilience: Circuit Breaker, Retry.",
        "Distributed Tracing: Sleuth, Zipkin.",
        "Event-Driven: Kafka, RabbitMQ."
    ]),
    ("Spring Boot Actuator & Monitoring", [
        "Actuator: health, metrics, env, info.",
        "Micrometer with Prometheus, Grafana.",
        "Custom Health Checks.",
        "Logging & Tracing: ELK stack."
    ]),
    ("Spring Boot Testing", [
        "Unit Testing: JUnit 5, Mockito, AssertJ.",
        "Integration Testing: @SpringBootTest, @MockBean.",
        "TestContainers for realistic testing.",
        "BDD with Cucumber."
    ]),
    ("Spring Boot Batch Processing", [
        "Spring Batch setup and job scheduling.",
        "Chunk-Oriented Processing.",
        "Tasklets and Scheduling.",
        "Retry & Skip Logic."
    ]),
    ("Spring Boot Cloud-Native Deployment", [
        "Dockerizing Spring Boot apps.",
        "Kubernetes Deployment: ConfigMaps, Secrets.",
        "CI/CD with Jenkins, GitLab CI.",
        "Cloud Deployments: AWS, Azure, GCP."
    ]),
    ("Spring Boot with Messaging Systems", [
        "Kafka integration for event streaming.",
        "RabbitMQ for queues and pub/sub.",
        "Spring Integration for messaging patterns."
    ]),
    ("Spring Boot Caching", [
        "Spring Cache with EhCache, Hazelcast, Redis.",
        "Custom eviction, management strategies.",
        "Distributed Caching."
    ]),
    ("Spring Boot Performance Optimization", [
        "JVM Tuning: memory, GC, VisualVM.",
        "Asynchronous with @Async, ExecutorService.",
        "Database tuning and indexing.",
        "Profiling and Bottleneck Analysis."
    ])
]

# Flatten the data for DataFrame
excel_data = [(topic, subtopic) for topic, subtopics in topics_data for subtopic in subtopics]

# Create DataFrame
df = pd.DataFrame(excel_data, columns=["Topic", "Details"])

# Use the current working directory
script_dir = os.getcwd()
file_path = os.path.join(script_dir, "spring_boot_topics.xlsx")

# Write the DataFrame to Excel
df.to_excel(file_path, index=False)

file_path
