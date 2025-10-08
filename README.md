# Microsoft Teams Transcript ELT Pipeline

## Table of Contents
- [Introduction](#introduction)
- [Objectives](#objectives)
  - [Business Objective](#business-objective)
  - [Technical Objective](#technical-objective)
- [Architecture Overview](#architecture-overview)
  - [Architecture Layers](#architecture-layers)
  - [End-to-End Flow](#end-to-end-flow)
- [Codebase Overview](#codebase-overview)
  - [Folder Structure](#folder-structure)
- [Metadata & SQL Schema](#metadata--sql-schema)
- [Azure AD & Teams Configuration](#azure-ad--teams-configuration)
- [Security & Access Control](#security--access-control)
- [Reliability, Scalability & Monitoring](#reliability-scalability--monitoring)
- [ADLS Folder Structure](#adls-folder-structure)
- [Non-Functional Requirements](#non-functional-requirements)
- [Risks & Mitigations](#risks--mitigations)
- [Setup & Deployment Instructions](#setup--deployment-instructions)
- [Summary](#summary)

## Introduction
This project implements an end-to-end ELT (Extract → Load → Transform) pipeline to automate collection, storage, and transformation of Microsoft Teams meeting transcripts for MobileLIVE.  
It ensures automated compliance with meeting transcript retention policies, centralized storage in Azure Data Lake for intelligence and audits, and AI-driven generation of structured meeting notes.

## Objectives

### Business Objective
- Automate Teams transcript collection and storage
- Centralize meeting intelligence for auditing and insights
- Generate AI-based meeting summaries

### Technical Objective
- Securely extract meeting and transcript data from Microsoft Graph API
- Store metadata in Azure SQL Database
- Store raw and AI-generated notes in Azure Data Lake Storage (ADLS)
- Transform transcripts into structured AI-generated summaries using Azure OpenAI

## Architecture Overview

### Architecture Layers

| Layer | Description | Key Azure Components |
|-------|-------------|----------------------|
| Extraction | Retrieve Teams metadata & transcripts | Graph API, Azure Function Apps, Azure SQL Database |
| Loading | Store metadata & transcripts | Azure SQL Database, Azure Blob Storage (teams-meeting-transcripts) |
| Transformation | Generate AI meeting notes | Azure Function App (AI Notes Generator), Azure OpenAI, Azure Blob Storage (ainotes) |
| Supporting | Security, observability, cross-service auth | Managed Identity, Azure Monitor, Application Insights |

### End-to-End Flow

#### 1. Extraction Layer
- Authentication: OAuth 2.0 Client Credentials Flow with Microsoft Graph
- Meeting Identification: Fetch official Teams Meeting ID via Graph API
- Transcript Extraction: List & download meeting transcripts
- Data Storage: Metadata to Azure SQL, transcript file to ADLS

#### 2. Loading Layer
- Timer-triggered function (every 15 mins): Checks and fetches pending transcripts, saves to ADLS, updates SQL

#### 3. Transformation Layer
- Function triggered on new transcript in ADLS, calls Azure OpenAI for summarization, updates SQL with AI notes

## Codebase Overview

| Current Name | Suggested New Name | Description |
|--------------|--------------------|-------------|
| meeting-transcriptfetch-scheduler | teams-transcript-fetcher-func | Periodically fetches Teams meeting transcripts from Graph API |
| meetingextractfunctionapp | teams-meeting-metadata-extractor-func | Extracts meeting metadata via Graph API |
| meetings-eventgrid-ainotes | teams-ainotes-generator-func | Generates AI meeting notes using OpenAI |
| Teamsmeeting-renewal-functionapp | teams-subscription-renewal-func | Renews Graph API event subscriptions |
| Hostingusers.py | teams-hosting-users-func | Populates SQL table with licensed Teams hosts |

### Folder Structure
src/
├─ teams-transcript-fetcher-func.py
├─ teams-meeting-metadata-extractor-func.py
├─ teams-ainotes-generator-func.py
├─ teams-subscription-renewal-func.py
└─ teams-hosting-users-func.py
docs/
└─ architecture.md


## Metadata & SQL Schema

**Table:** TeamsMeetingMetadata

| Column | Type | Description |
|--------|------|-------------|
| UniqueId (PK) | VARCHAR(100) | Primary Key |
| TeamsMeetingId | VARCHAR(200) | Teams Meeting Identifier |
| OrganizerEmail | VARCHAR | Organizer Email |
| OrganizerObjectId | VARCHAR | Organizer Object ID |
| Subject, StartTime, EndTime, JoinUrl | Varchar/Datetime | Meeting details |
| TranscriptUrl | VARCHAR | URL to transcript in ADLS |
| AdlsPath | VARCHAR | Raw transcript path in ADLS |
| AINotesPath | VARCHAR | AI-generated notes path |
| Notes | Text | AI-generated text notes |
| Status | VARCHAR | (PENDING/IN_PROGRESS/COMPLETED) |
| MeetingCompletionStatus | VARCHAR | (SUCCESS/FAILURE) |
| TranscriptStatus | VARCHAR | |
| CreatedAt, LastUpdatedAt | Datetime | Metadata audit fields |
| IsParent | BIT | Recurring meeting flag |

## Azure AD & Teams Configuration

- Graph API: Application & Delegated permissions (see docs)
- Teams Meeting Policy: Enable transcription & recording
- Teams Application Access Policy: Allow app to read meeting artifacts

## Security & Access Control

- OAuth 2.0 Client Credentials; secrets in Azure Key Vault
- Managed Identity for Functions
- Private endpoints for SQL & ADLS
- TLS enforced, encryption at rest/in-transit
- 90-day lifecycle management

## Reliability, Scalability & Monitoring

| Category | Design Choice |
|----------|---------------|
| Scalability | Azure Function Consumption Plan (auto-scale) |
| Retry Policy | Exponential backoff for 429/503 |
| Monitoring | Application Insights (logs/metrics); Azure Monitor (alerts) |
| Durability | SQL + ADLS persistent storage |
| Alerting | Azure Monitor alerts |

## ADLS Folder Structure

- **Raw Transcripts:**  
  `teams-meeting-transcripts/<organizer>/<subject>/<date>/transcript.txt`
- **AI Notes:**  
  `teams-meeting-ainotes/<organizer>/<subject>/<date>/ainotes.txt`

## Non-Functional Requirements

- Availability: 99.9%
- Latency: < 5 mins transcript ingestion
- Retention: 7 years in ADLS
- Security Compliance: SOC 2 / ISO 27001
- Cost Optimization: Archive after 90 days

## Risks & Mitigations

| Risk | Mitigation |
|------|------------|
| API Rate Limit | Exponential retry |
| Token Expiry | Auto-refresh via Key Vault |
| Duplicate Triggers | Validate SQL state before process |
| Graph Policy Delay | Wait 24 hrs post policy assignment |
| Transcript Unavailability | Retry next execution |

## Summary

This serverless ELT pipeline automates extraction, storage, and transformation of Microsoft Teams transcripts, integrating with Azure SQL, ADLS, and Azure OpenAI.  
