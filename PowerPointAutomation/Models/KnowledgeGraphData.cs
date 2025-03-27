using System;
using System.Collections.Generic;
using System.Drawing;

namespace PowerPointAutomation.Models
{
    /// <summary>
    /// Contains sample data structures for knowledge graph examples
    /// This class provides pre-configured data for demonstration purposes
    /// </summary>
    public static class KnowledgeGraphData
    {
        /// <summary>
        /// Gets a collection of sample slide content for a complete knowledge graph presentation
        /// </summary>
        /// <returns>A dictionary of slide content objects keyed by slide index</returns>
        public static Dictionary<int, SlideContent> GetSamplePresentation()
        {
            var slides = new Dictionary<int, SlideContent>();

            // Title slide
            slides.Add(1, new TitleSlideContent(
                "Knowledge Graphs",
                "A Comprehensive Introduction",
                "PowerPoint Automation Demo",
                "The title slide introduces the presentation topic with a visually appealing layout. " +
                "The main title uses a larger font with the primary color, while the subtitle uses a " +
                "slightly smaller font with the secondary color. This establishes the visual hierarchy " +
                "that will be consistent throughout the presentation."
            ));

            // Introduction slide
            slides.Add(2, new BulletSlideContent(
                "Introduction to Knowledge Graphs",
                new string[] {
                    "Knowledge graphs represent information as interconnected entities and relationships",
                    "Semantic networks that represent real-world entities (objects, events, concepts)",
                    "Bridge structured and unstructured data for human and machine interpretation",
                    "Enable sophisticated reasoning, discovery, and analysis capabilities",
                    "Create a flexible yet robust foundation for knowledge management"
                },
                "This slide introduces the fundamental concept of knowledge graphs as networks of " +
                "entities and relationships. Emphasize how they differ from traditional data structures " +
                "by explicitly modeling connections."
            ));

            // Core components slide
            slides.Add(3, new BulletSlideContent(
                "Core Components of Knowledge Graphs",
                new string[] {
                    "Nodes (Entities): Discrete objects, concepts, events, or states",
                    "• Unique identifiers, categorized by type, contain properties",
                    "Edges (Relationships): Connect nodes and define how entities relate",
                    "• Directed connections with semantic meaning, typed, may contain properties",
                    "Labels and Properties: Provide additional context and attributes",
                    "• Node labels denote entity types, edge labels specify relationship types"
                },
                "This slide outlines the three fundamental building blocks of knowledge graphs. " +
                "The nested bullet points provide more detail about each component."
            ));

            // Structural example slide
            slides.Add(4, new DiagramSlideContent(
                "Structural Example",
                "A simple knowledge graph fragment representing company information",
                DiagramType.KnowledgeGraph,
                "This slide presents a visual example of a knowledge graph structure. " +
                "The diagram shows how entities are connected through relationships, " +
                "with each having specific properties."
            ));

            // Theoretical foundations slide
            slides.Add(5, new TwoColumnSlideContent(
                "Theoretical Foundations",
                new string[] {
                    "Graph Theory",
                    "• Connectivity, centrality, community structure",
                    "• Path analysis, network algorithms",
                    "Semantic Networks",
                    "• Conceptual associations, hierarchical organizations",
                    "• Meaning representation"
                },
                new string[] {
                    "Ontological Modeling",
                    "• Class hierarchies, property definitions",
                    "• Axioms and rules, domain modeling",
                    "Knowledge Representation",
                    "• First-order logic, description logics",
                    "• Frame systems, semantic triples"
                },
                "This slide presents the theoretical foundations that knowledge graphs build upon. " +
                "The two-column layout helps organize related but distinct concepts."
            ));

            // Implementation technologies slide
            slides.Add(6, new BulletSlideContent(
                "Implementation Technologies",
                new string[] {
                    "Data Models",
                    "• RDF (Resource Description Framework)",
                    "• Property Graphs",
                    "• Hypergraphs",
                    "• Knowledge Graph Embeddings",
                    "Storage Solutions",
                    "• Native Graph Databases (Neo4j, TigerGraph)",
                    "• RDF Triple Stores (AllegroGraph, Stardog)",
                    "• Multi-Model Databases (ArangoDB, OrientDB)"
                },
                "This slide covers the various technologies used to implement knowledge graphs. " +
                "It presents both data models and storage solutions."
            ));

            // Construction approaches slide
            slides.Add(7, new TwoColumnSlideContent(
                "Construction Approaches",
                new string[] {
                    "Manual Curation",
                    "• Expert-driven construction ensuring high quality",
                    "• Time-intensive, difficult to scale",
                    "• Critical domains requiring accuracy (healthcare, legal)",
                    "Automated Extraction",
                    "• Information extraction from text",
                    "• Wrapper induction from web pages",
                    "• Database transformation from relational data"
                },
                new string[] {
                    "Hybrid Approaches",
                    "• Bootstrap and refine: Automated with manual verification",
                    "• Pattern-based expansion: Using patterns to extend examples",
                    "• Distant supervision: Leveraging existing knowledge",
                    "• Continuous feedback: Incorporating user corrections",
                    "Evaluation Criteria",
                    "• Accuracy, coverage, consistency",
                    "• Semantic validity, alignment with domain knowledge"
                },
                "This slide presents different methodologies for constructing knowledge graphs. " +
                "The two-column layout creates a natural comparison between approaches."
            ));

            // Machine learning integration slide
            slides.Add(8, new DiagramSlideContent(
                "Machine Learning Integration",
                "How knowledge graphs and machine learning interact",
                DiagramType.MLIntegration,
                "This slide visualizes the bidirectional relationship between knowledge graphs " +
                "and machine learning. The circular diagram shows how machine learning can help " +
                "build and enhance knowledge graphs, while knowledge graphs can improve machine " +
                "learning models through structured knowledge."
            ));

            // Applications slide
            slides.Add(9, new BulletSlideContent(
                "Applications & Use Cases",
                new string[] {
                    "Enterprise Knowledge Management",
                    "• Corporate memory, expertise location, document management",
                    "Search and Recommendation Systems",
                    "• Semantic search, context-aware recommendations, knowledge panels",
                    "Research and Discovery",
                    "• Scientific literature analysis, drug discovery, patent analysis",
                    "Customer Intelligence",
                    "• 360° customer view, journey mapping, nuanced segmentation",
                    "Compliance and Risk Management",
                    "• Regulatory compliance, fraud detection, anti-money laundering"
                },
                "This slide showcases diverse applications of knowledge graphs across domains. " +
                "The hierarchical structure organizes use cases by industry or function."
            ));

            // Advantages and challenges slide
            slides.Add(10, new TwoColumnSlideContent(
                "Advantages & Challenges",
                new string[] {
                    "Key Advantages",
                    "• Contextual Understanding: Data with semantic context",
                    "• Flexibility: Adaptable to evolving information needs",
                    "• Integration Capability: Unifies diverse data sources",
                    "• Inferential Power: Discovers implicit knowledge",
                    "• Human-Interpretable: Aligns with conceptual understanding"
                },
                new string[] {
                    "Implementation Challenges",
                    "• Construction Complexity: Significant effort required",
                    "• Schema Evolution: Maintaining consistency while growing",
                    "• Performance at Scale: Optimizing for large graphs",
                    "• Quality Assurance: Ensuring accuracy across assertions",
                    "• User Adoption: Requiring new query paradigms"
                },
                "This slide presents a balanced view of both the advantages and challenges of " +
                "knowledge graph implementations. The side-by-side comparison helps decision-makers " +
                "understand both the benefits and potential obstacles."
            ));

            // Future directions slide
            slides.Add(11, new BulletSlideContent(
                "Future Directions",
                new string[] {
                    "Self-Improving Knowledge Graphs",
                    "• Automated knowledge acquisition and contradiction detection",
                    "• Confidence scoring and active learning",
                    "Multimodal Knowledge Graphs",
                    "• Visual, temporal, spatial, and numerical integration",
                    "• Cross-modal reasoning and representation",
                    "Neuro-Symbolic Integration",
                    "• Combining neural networks with symbolic logic",
                    "• Using knowledge graphs to explain AI decisions",
                    "• Foundation model integration with knowledge graphs"
                },
                "This slide explores emerging trends and future developments in knowledge graph " +
                "technology. The hierarchical structure helps organize related concepts."
            ));

            // Conclusion slide
            slides.Add(12, new ConclusionSlideContent(
                "Conclusion",
                "Knowledge graphs represent a transformative approach to information management, enabling organizations to move beyond data silos toward connected intelligence. By explicitly modeling relationships between entities, knowledge graphs provide context that traditional databases lack, supporting sophisticated reasoning and discovery.\n\n" +
                "While implementing knowledge graphs presents challenges in construction, maintenance, and scalability, the benefits of contextual understanding, flexible integration, and inferential capabilities make them increasingly essential for organizations dealing with complex, interconnected information.",
                "Thank You!",
                "contact@example.com",
                "This conclusion slide summarizes the key takeaways about knowledge graphs. " +
                "It reinforces the main value proposition while acknowledging the implementation " +
                "challenges."
            ));

            return slides;
        }

        /// <summary>
        /// Returns a sample knowledge graph structure with entities, relationships, and properties
        /// </summary>
        /// <returns>A KnowledgeGraph object with sample data</returns>
        public static KnowledgeGraph GetSampleKnowledgeGraph()
        {
            // Create a sample knowledge graph
            var graph = new KnowledgeGraph();

            // Create entities
            var company = new Entity("TechCorp", "Company");
            company.Properties.Add("founded", "2010");
            company.Properties.Add("revenue", "$2.5M");
            company.Properties.Add("location", "San Francisco");
            graph.AddEntity(company);

            var person1 = new Entity("John Doe", "Person");
            person1.Properties.Add("role", "Engineer");
            person1.Properties.Add("expertise", "AI, Python");
            person1.Properties.Add("joined", "2015-03-10");
            graph.AddEntity(person1);

            var person2 = new Entity("Jane Smith", "Person");
            person2.Properties.Add("role", "Manager");
            person2.Properties.Add("expertise", "Leadership, Knowledge Graphs");
            person2.Properties.Add("joined", "2012-06-22");
            graph.AddEntity(person2);

            var product = new Entity("ProductX", "Product");
            product.Properties.Add("launchDate", "2020-01-15");
            product.Properties.Add("category", "Software");
            product.Properties.Add("version", "1.2.3");
            graph.AddEntity(product);

            var feature1 = new Entity("AI Assistant", "Feature");
            feature1.Properties.Add("status", "Released");
            feature1.Properties.Add("complexity", "High");
            graph.AddEntity(feature1);

            var feature2 = new Entity("Data Visualization", "Feature");
            feature2.Properties.Add("status", "In Development");
            feature2.Properties.Add("complexity", "Medium");
            graph.AddEntity(feature2);

            // Create relationships
            graph.AddRelationship(company, person1, "EMPLOYS");
            graph.AddRelationship(company, person2, "EMPLOYS");
            graph.AddRelationship(person2, person1, "MANAGES");
            graph.AddRelationship(company, product, "PRODUCES");
            graph.AddRelationship(product, feature1, "HAS_FEATURE");
            graph.AddRelationship(product, feature2, "HAS_FEATURE");
            graph.AddRelationship(person1, feature1, "DEVELOPS");
            graph.AddRelationship(person1, feature2, "DEVELOPS");
            graph.AddRelationship(person2, product, "OVERSEES");

            return graph;
        }
    }

    /// <summary>
    /// Represents a knowledge graph with entities and relationships
    /// </summary>
    public class KnowledgeGraph
    {
        /// <summary>
        /// Collection of entities in the graph
        /// </summary>
        public List<Entity> Entities { get; } = new List<Entity>();

        /// <summary>
        /// Collection of relationships in the graph
        /// </summary>
        public List<Relationship> Relationships { get; } = new List<Relationship>();

        /// <summary>
        /// Adds an entity to the graph
        /// </summary>
        /// <param name="entity">The entity to add</param>
        public void AddEntity(Entity entity)
        {
            Entities.Add(entity);
        }

        /// <summary>
        /// Creates and adds a relationship between two entities
        /// </summary>
        /// <param name="source">Source entity</param>
        /// <param name="target">Target entity</param>
        /// <param name="type">Relationship type</param>
        /// <returns>The created relationship</returns>
        public Relationship AddRelationship(Entity source, Entity target, string type)
        {
            var relationship = new Relationship(source, target, type);
            Relationships.Add(relationship);
            return relationship;
        }

        /// <summary>
        /// Gets all relationships for a specific entity
        /// </summary>
        /// <param name="entity">The entity to find relationships for</param>
        /// <returns>List of relationships</returns>
        public List<Relationship> GetRelationshipsForEntity(Entity entity)
        {
            return Relationships.FindAll(r => r.Source == entity || r.Target == entity);
        }

        /// <summary>
        /// Gets all entities of a specific type
        /// </summary>
        /// <param name="type">Entity type to filter by</param>
        /// <returns>List of matching entities</returns>
        public List<Entity> GetEntitiesByType(string type)
        {
            return Entities.FindAll(e => e.Type == type);
        }
    }

    /// <summary>
    /// Represents an entity in a knowledge graph
    /// </summary>
    public class Entity
    {
        /// <summary>
        /// Unique identifier for the entity
        /// </summary>
        public string Id { get; } = Guid.NewGuid().ToString();

        /// <summary>
        /// Name/label of the entity
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Type/category of the entity
        /// </summary>
        public string Type { get; set; }

        /// <summary>
        /// Properties/attributes of the entity
        /// </summary>
        public Dictionary<string, string> Properties { get; } = new Dictionary<string, string>();

        /// <summary>
        /// Default constructor
        /// </summary>
        public Entity() { }

        /// <summary>
        /// Constructor with name and type
        /// </summary>
        /// <param name="name">Entity name</param>
        /// <param name="type">Entity type</param>
        public Entity(string name, string type)
        {
            Name = name;
            Type = type;
        }

        /// <summary>
        /// Returns a string representation of the entity
        /// </summary>
        /// <returns>String representation</returns>
        public override string ToString()
        {
            return $"{Name} ({Type})";
        }
    }

    /// <summary>
    /// Represents a relationship between two entities in a knowledge graph
    /// </summary>
    public class Relationship
    {
        /// <summary>
        /// Unique identifier for the relationship
        /// </summary>
        public string Id { get; } = Guid.NewGuid().ToString();

        /// <summary>
        /// Source entity (start of the relationship)
        /// </summary>
        public Entity Source { get; set; }

        /// <summary>
        /// Target entity (end of the relationship)
        /// </summary>
        public Entity Target { get; set; }

        /// <summary>
        /// Type/label of the relationship
        /// </summary>
        public string Type { get; set; }

        /// <summary>
        /// Properties/attributes of the relationship
        /// </summary>
        public Dictionary<string, string> Properties { get; } = new Dictionary<string, string>();

        /// <summary>
        /// Default constructor
        /// </summary>
        public Relationship() { }

        /// <summary>
        /// Constructor with source, target, and type
        /// </summary>
        /// <param name="source">Source entity</param>
        /// <param name="target">Target entity</param>
        /// <param name="type">Relationship type</param>
        public Relationship(Entity source, Entity target, string type)
        {
            Source = source;
            Target = target;
            Type = type;
        }

        /// <summary>
        /// Returns a string representation of the relationship
        /// </summary>
        /// <returns>String representation</returns>
        public override string ToString()
        {
            return $"{Source.Name} -{Type}-> {Target.Name}";
        }
    }
}