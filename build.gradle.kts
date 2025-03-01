plugins {
    kotlin("jvm") version "2.1.10"
}

group = "io.luxus.dexcel"
version = "1.0-SNAPSHOT"

repositories {
    mavenCentral()
}

dependencies {
    implementation("org.apache.poi:poi:5.4.0")
    implementation("org.apache.poi:poi-ooxml:5.4.0")

    testImplementation(kotlin("test"))
    testImplementation("io.kotest:kotest-runner-junit5:5.6.2")
    testImplementation("io.kotest:kotest-assertions-core:5.6.2")
    testImplementation("io.mockk:mockk:1.13.5")
}

tasks.test {
    useJUnitPlatform()
}
kotlin {
    jvmToolchain(8)
}
