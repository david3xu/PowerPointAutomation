using System;
using System.Collections.Generic;
using Microsoft.Office.Core;
using PowerPointShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PowerPointShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAutomation.Utilities
{
    /// <summary>
    /// Helper class for creating complex PowerPoint animations
    /// </summary>
    /// <remarks>
    /// This class provides methods to simplify the creation of animations in PowerPoint.
    /// It abstracts common animation patterns and provides a more intuitive interface
    /// for working with PowerPoint's animation model.
    /// </remarks>
    public static class AnimationHelper
    {
        #region Simple Animation Methods

        /// <summary>
        /// Creates a sequential fade-in animation for multiple shapes
        /// </summary>
        /// <param name="slide">The slide containing the shapes</param>
        /// <param name="shapes">Array of shapes to animate</param>
        /// <param name="clickToStart">Whether the first shape should animate on click</param>
        /// <param name="duration">Duration of each animation in seconds</param>
        /// <param name="delay">Delay between animations in seconds</param>
        /// <returns>Array of created animation effects</returns>
        public static Effect[] CreateSequentialFadeAnimation(
            Microsoft.Office.Interop.PowerPoint.Slide slide,
            PowerPointShape[] shapes,
            bool clickToStart = true,
            float duration = 0.5f,
            float delay = 0.2f)
        {
            if (shapes == null || shapes.Length == 0)
                return new Effect[0];

            Effect[] effects = new Effect[shapes.Length];

            // Add the first shape with appropriate trigger
            MsoAnimTriggerType firstTrigger = clickToStart ?
                MsoAnimTriggerType.msoAnimTriggerOnPageClick :
                MsoAnimTriggerType.msoAnimTriggerWithPrevious;

            effects[0] = slide.TimeLine.MainSequence.AddEffect(
                shapes[0],
                MsoAnimEffect.msoAnimEffectFade,
                MsoAnimateByLevel.msoAnimateLevelNone,
                firstTrigger);

            effects[0].Timing.Duration = duration;

            // Add remaining shapes to animate after the previous one
            for (int i = 1; i < shapes.Length; i++)
            {
                effects[i] = slide.TimeLine.MainSequence.AddEffect(
                    shapes[i],
                    MsoAnimEffect.msoAnimEffectFade,
                    MsoAnimateByLevel.msoAnimateLevelNone,
                    MsoAnimTriggerType.msoAnimTriggerAfterPrevious);

                effects[i].Timing.Duration = duration;

                // Set delay for better visual flow
                if (delay > 0)
                {
                    effects[i].Timing.TriggerDelayTime = delay;
                }
            }

            return effects;
        }

        /// <summary>
        /// Creates a build animation for bullet points in a text shape
        /// </summary>
        /// <param name="slide">The slide containing the shape</param>
        /// <param name="textShape">The text shape with bullet points</param>
        /// <param name="clickToStart">Whether to start on click</param>
        /// <param name="duration">Duration of each animation</param>
        /// <returns>The created animation effect</returns>
        public static Effect CreateBulletPointAnimation(
            Microsoft.Office.Interop.PowerPoint.Slide slide,
            PowerPointShape textShape,
            bool clickToStart = true,
            float duration = 0.3f)
        {
            MsoAnimTriggerType trigger = clickToStart ?
                MsoAnimTriggerType.msoAnimTriggerOnPageClick :
                MsoAnimTriggerType.msoAnimTriggerWithPrevious;

            // For now, animate the whole shape since paragraph level might not be supported
            Effect effect = slide.TimeLine.MainSequence.AddEffect(
                textShape,
                MsoAnimEffect.msoAnimEffectFade,
                MsoAnimateByLevel.msoAnimateLevelNone,
                trigger);

            effect.Timing.Duration = duration;

            return effect;
        }

        /// <summary>
        /// Creates an emphasis animation for a shape
        /// </summary>
        /// <param name="slide">The slide containing the shape</param>
        /// <param name="shape">The shape to animate</param>
        /// <param name="effect">The animation effect to use</param>
        /// <param name="clickToStart">Whether to start on click</param>
        /// <param name="duration">Duration of the animation</param>
        /// <returns>The created animation effect</returns>
        public static Effect CreateEmphasisAnimation(
            Microsoft.Office.Interop.PowerPoint.Slide slide,
            PowerPointShape shape,
            MsoAnimEffect effect = MsoAnimEffect.msoAnimEffectAppear,
            bool clickToStart = true,
            float duration = 0.7f)
        {
            MsoAnimTriggerType trigger = clickToStart ?
                MsoAnimTriggerType.msoAnimTriggerOnPageClick :
                MsoAnimTriggerType.msoAnimTriggerWithPrevious;

            Effect animEffect = slide.TimeLine.MainSequence.AddEffect(
                shape,
                effect,
                MsoAnimateByLevel.msoAnimateLevelNone,
                trigger);

            animEffect.Timing.Duration = duration;
            
            // Don't set read-only properties
            
            return animEffect;
        }

        /// <summary>
        /// Creates a path animation for a shape
        /// </summary>
        /// <param name="slide">The slide containing the shape</param>
        /// <param name="shape">The shape to animate</param>
        /// <param name="fromX">Starting X coordinate</param>
        /// <param name="fromY">Starting Y coordinate</param>
        /// <param name="toX">Ending X coordinate</param>
        /// <param name="toY">Ending Y coordinate</param>
        /// <param name="clickToStart">Whether to start on click</param>
        /// <param name="duration">Duration of the animation</param>
        /// <returns>The created animation effect</returns>
        public static Effect CreatePathAnimation(
            Microsoft.Office.Interop.PowerPoint.Slide slide,
            PowerPointShape shape,
            float fromX,
            float fromY,
            float toX,
            float toY,
            bool clickToStart = true,
            float duration = 1.0f)
        {
            MsoAnimTriggerType trigger = clickToStart ?
                MsoAnimTriggerType.msoAnimTriggerOnPageClick :
                MsoAnimTriggerType.msoAnimTriggerWithPrevious;

            // Move the shape to the starting position
            shape.Left = fromX;
            shape.Top = fromY;

            // Add a basic animation instead of a path
            Effect effect = slide.TimeLine.MainSequence.AddEffect(
                shape,
                MsoAnimEffect.msoAnimEffectFade,
                MsoAnimateByLevel.msoAnimateLevelNone,
                trigger);

            effect.Timing.Duration = duration;

            return effect;
        }

        #endregion

        #region Complex Animation Sequences

        /// <summary>
        /// Creates a "one by one" animation sequence where items appear and move to their final positions
        /// </summary>
        /// <param name="slide">The slide containing the shapes</param>
        /// <param name="shapes">Array of shapes to animate</param>
        /// <param name="finalPositions">Array of final positions (left, top) for each shape</param>
        /// <param name="startPosition">Starting position for all shapes</param>
        /// <param name="clickToStart">Whether the first shape should animate on click</param>
        /// <returns>Array of created animation effects</returns>
        public static Effect[] CreateOneByOneAppearanceSequence(
            Microsoft.Office.Interop.PowerPoint.Slide slide,
            PowerPointShape[] shapes,
            (float Left, float Top)[] finalPositions,
            (float Left, float Top) startPosition,
            bool clickToStart = true)
        {
            if (shapes == null || shapes.Length == 0 || finalPositions.Length != shapes.Length)
                return new Effect[0];

            List<Effect> effects = new List<Effect>();

            // First shape animation trigger
            MsoAnimTriggerType firstTrigger = clickToStart ?
                MsoAnimTriggerType.msoAnimTriggerOnPageClick :
                MsoAnimTriggerType.msoAnimTriggerWithPrevious;

            // For each shape, initially position it at start, then animate to final position
            for (int i = 0; i < shapes.Length; i++)
            {
                // Set initial position
                shapes[i].Left = startPosition.Left;
                shapes[i].Top = startPosition.Top;

                // First make it appear
                Effect appearEffect = slide.TimeLine.MainSequence.AddEffect(
                    shapes[i],
                    MsoAnimEffect.msoAnimEffectFade,
                    MsoAnimateByLevel.msoAnimateLevelNone,
                    i == 0 ? firstTrigger : MsoAnimTriggerType.msoAnimTriggerAfterPrevious);

                appearEffect.Timing.Duration = 0.3f;
                effects.Add(appearEffect);

                // Then move it to final position
                Effect moveEffect = slide.TimeLine.MainSequence.AddEffect(
                    shapes[i],
                    MsoAnimEffect.msoAnimEffectPath4PointStar, // Using a custom path
                    MsoAnimateByLevel.msoAnimateLevelNone,
                    MsoAnimTriggerType.msoAnimTriggerWithPrevious);

                // Calculate movement path
                float deltaX = finalPositions[i].Left - startPosition.Left;
                float deltaY = finalPositions[i].Top - startPosition.Top;

                moveEffect.Timing.Duration = 0.5f;
                effects.Add(moveEffect);

                // Add a delay before the next item
                if (i < shapes.Length - 1)
                {
                    moveEffect.Timing.TriggerDelayTime = 0.1f;
                }
            }

            return effects.ToArray();
        }

        /// <summary>
        /// Creates a flourish animation for a title or important element
        /// </summary>
        /// <param name="slide">The slide containing the shape</param>
        /// <param name="shape">The shape to animate</param>
        /// <param name="clickToStart">Whether to start on click</param>
        /// <returns>Array of created animation effects</returns>
        public static Effect[] CreateFlourishAnimation(
            Microsoft.Office.Interop.PowerPoint.Slide slide,
            PowerPointShape shape,
            bool clickToStart = true)
        {
            List<Effect> effects = new List<Effect>();

            // First animation trigger
            MsoAnimTriggerType trigger = clickToStart ?
                MsoAnimTriggerType.msoAnimTriggerOnPageClick :
                MsoAnimTriggerType.msoAnimTriggerWithPrevious;

            // First, fade in
            Effect fadeEffect = slide.TimeLine.MainSequence.AddEffect(
                shape,
                MsoAnimEffect.msoAnimEffectFade,
                MsoAnimateByLevel.msoAnimateLevelNone,
                trigger);

            fadeEffect.Timing.Duration = 0.5f;
            effects.Add(fadeEffect);

            // Then, add a subtle grow effect
            Effect growEffect = slide.TimeLine.MainSequence.AddEffect(
                shape,
                MsoAnimEffect.msoAnimEffectGrowAndTurn,
                MsoAnimateByLevel.msoAnimateLevelNone,
                MsoAnimTriggerType.msoAnimTriggerWithPrevious);

            growEffect.Timing.Duration = 0.7f;
            effects.Add(growEffect);

            // Finally, add a subtle emphasis
            Effect emphasisEffect = slide.TimeLine.MainSequence.AddEffect(
                shape,
                MsoAnimEffect.msoAnimEffectTeeter,
                MsoAnimateByLevel.msoAnimateLevelNone,
                MsoAnimTriggerType.msoAnimTriggerAfterPrevious);

            emphasisEffect.Timing.Duration = 0.5f;
            effects.Add(emphasisEffect);

            return effects.ToArray();
        }

        /// <summary>
        /// Creates a diagram build animation that reveals components one by one
        /// </summary>
        /// <param name="slide">The slide containing the shapes</param>
        /// <param name="backgroundShape">The background shape for the diagram (if any)</param>
        /// <param name="nodeShapes">Array of node shapes</param>
        /// <param name="connectionShapes">Array of connection shapes</param>
        /// <param name="clickToStart">Whether to start on click</param>
        /// <returns>Array of created animation effects</returns>
        public static Effect[] CreateDiagramBuildAnimation(
            Microsoft.Office.Interop.PowerPoint.Slide slide,
            PowerPointShape backgroundShape,
            PowerPointShape[] nodeShapes,
            PowerPointShape[] connectionShapes,
            bool clickToStart = true)
        {
            List<Effect> effects = new List<Effect>();

            // First animation trigger
            MsoAnimTriggerType trigger = clickToStart ?
                MsoAnimTriggerType.msoAnimTriggerOnPageClick :
                MsoAnimTriggerType.msoAnimTriggerWithPrevious;

            // First, animate the background if provided
            if (backgroundShape != null)
            {
                Effect bgEffect = slide.TimeLine.MainSequence.AddEffect(
                    backgroundShape,
                    MsoAnimEffect.msoAnimEffectFade,
                    MsoAnimateByLevel.msoAnimateLevelNone,
                    trigger);

                bgEffect.Timing.Duration = 0.5f;
                effects.Add(bgEffect);

                // Next animations start after the background
                trigger = MsoAnimTriggerType.msoAnimTriggerAfterPrevious;
            }

            // Then, animate nodes one by one with different effects based on position
            if (nodeShapes != null && nodeShapes.Length > 0)
            {
                for (int i = 0; i < nodeShapes.Length; i++)
                {
                    // Select different animation effects based on position
                    MsoAnimDirection direction = GetDirectionForIndex(i);

                    Effect nodeEffect = slide.TimeLine.MainSequence.AddEffect(
                        nodeShapes[i],
                        MsoAnimEffect.msoAnimEffectFly,
                        MsoAnimateByLevel.msoAnimateLevelNone,
                        i == 0 ? trigger : MsoAnimTriggerType.msoAnimTriggerAfterPrevious);

                    nodeEffect.EffectParameters.Direction = direction;
                    nodeEffect.Timing.Duration = 0.5f;
                    effects.Add(nodeEffect);
                }
            }

            // Finally, animate connections
            if (connectionShapes != null && connectionShapes.Length > 0)
            {
                for (int i = 0; i < connectionShapes.Length; i++)
                {
                    Effect connEffect = slide.TimeLine.MainSequence.AddEffect(
                        connectionShapes[i],
                        MsoAnimEffect.msoAnimEffectWipe,
                        MsoAnimateByLevel.msoAnimateLevelNone,
                        MsoAnimTriggerType.msoAnimTriggerAfterPrevious);

                    // Wipe direction based on connection orientation
                    connEffect.EffectParameters.Direction = MsoAnimDirection.msoAnimDirectionLeft;
                    connEffect.Timing.Duration = 0.4f;
                    effects.Add(connEffect);
                }
            }

            return effects.ToArray();
        }

        #endregion

        #region Helper Methods

        /// <summary>
        /// Gets a direction for animation based on an index
        /// </summary>
        /// <param name="index">The index to determine direction</param>
        /// <returns>Animation direction</returns>
        private static MsoAnimDirection GetDirectionForIndex(int index)
        {
            switch (index % 4)
            {
                case 0:
                    return MsoAnimDirection.msoAnimDirectionUp;
                case 1:
                    return MsoAnimDirection.msoAnimDirectionLeft;
                case 2:
                    return MsoAnimDirection.msoAnimDirectionRight;
                case 3:
                    return MsoAnimDirection.msoAnimDirectionDown;
                default:
                    return MsoAnimDirection.msoAnimDirectionUp;
            }
        }

        /// <summary>
        /// Groups existing animations into click groups where each group plays on a single click
        /// </summary>
        /// <param name="sequence">The animation sequence to modify</param>
        /// <param name="groupSize">Number of animations per click group</param>
        public static void GroupAnimationsIntoClickSteps(Sequence sequence, int groupSize)
        {
            if (sequence.Count <= 1 || groupSize <= 1)
                return;

            // Start with the second animation (leave first animation's trigger unchanged)
            for (int i = 1; i < sequence.Count; i++)
            {
                Effect effect = sequence[i + 1]; // PowerPoint uses 1-based indexing

                // If this should be the first animation in a new group, set it to OnClick
                if (i % groupSize == 0)
                {
                    effect.Timing.TriggerType = MsoAnimTriggerType.msoAnimTriggerOnPageClick;
                }
                else
                {
                    // Otherwise, set it to play with the previous animation
                    effect.Timing.TriggerType = MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                }
            }
        }

        /// <summary>
        /// Creates a text reveal animation that shows text one character, word, or line at a time
        /// </summary>
        /// <param name="slide">The slide containing the shape</param>
        /// <param name="textShape">The text shape to animate</param>
        /// <param name="revealBy">How to reveal the text</param>
        /// <param name="clickToStart">Whether to start on click</param>
        /// <param name="duration">Duration of the animation</param>
        /// <returns>The created animation effect</returns>
        public static Effect CreateTextRevealAnimation(
            Microsoft.Office.Interop.PowerPoint.Slide slide,
            PowerPointShape textShape,
            MsoAnimTextUnitEffect revealBy = MsoAnimTextUnitEffect.msoAnimTextUnitEffectByWord,
            bool clickToStart = true,
            float duration = 1.0f)
        {
            MsoAnimTriggerType trigger = clickToStart ?
                MsoAnimTriggerType.msoAnimTriggerOnPageClick :
                MsoAnimTriggerType.msoAnimTriggerWithPrevious;

            Effect effect = slide.TimeLine.MainSequence.AddEffect(
                textShape,
                MsoAnimEffect.msoAnimEffectAppear,
                MsoAnimateByLevel.msoAnimateLevelNone,
                trigger);

            // Set text animation properties
            effect.Timing.Duration = duration;
            // TODO: Fix this line - PowerPoint version compatibility issue
            // effect.Behaviors[1].SetEffect.TextUnitEffect = revealBy;

            return effect;
        }

        #endregion
    }
}
